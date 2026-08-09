//! Transactional snapshot mutation for persisted slide and master records.

use super::super::{
    AnimationInfo, BuildList, BuildListEntry, ExtendedTimeNode, parse_animation_info,
    parse_slide_animation_extension, write_animation_info, write_build_list,
    write_extended_time_node,
};
use super::semantic::{
    Editor, EditorLimits, LegacyShapeAnimation, PersistAnimation, Scope, Timeline,
};
use super::validation::{validate_extension, validate_limits};
use crate::consts::RecordType;
use crate::embedded::object::editor::Editor as ObjectEditor;
use crate::package::{Error, Result};
use crate::records::Record;
use std::collections::{BTreeSet, HashSet};

use super::semantic::{ESCHER_CLIENT_DATA, ESCHER_SP, ESCHER_SP_CONTAINER};

impl Editor {
    /// Opens an animation editor over the persisted records of a package.
    ///
    /// # Errors
    ///
    /// Returns an error if `limits` is invalid, the package or one of its
    /// persisted records is malformed, or the animation data exceeds the
    /// configured limits.
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
            let scope = if kind == RecordType::Slide.as_u16() {
                Scope::Slide
            } else if kind == RecordType::MainMaster.as_u16() {
                Scope::MainMaster
            } else {
                continue;
            };
            let (parsed, consumed) = Record::parse(&record, 0)?;
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

    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changed
    }

    #[must_use]
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

    #[must_use]
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

    #[must_use]
    pub fn legacy_shape_animations(&self) -> Vec<LegacyShapeAnimation> {
        self.entries
            .iter()
            .flat_map(|entry| entry.legacy.clone())
            .collect()
    }

    #[must_use]
    pub fn find_shape(&self, persist_id: u32, shape_id: u32) -> Option<LegacyShapeAnimation> {
        self.entries
            .iter()
            .find(|entry| entry.persist_id == persist_id)
            .and_then(|entry| entry.legacy.iter().find(|value| value.shape_id == shape_id))
            .cloned()
    }

    /// Adds a child to the root timing container at `index`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update(&mut self, persist_id: u32, index: usize, node: ExtendedTimeNode) -> Result<()> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        let root = entry
            .extension
            .time_node
            .as_mut()
            .ok_or_else(|| Error::InvalidFormat("timeline has no root node".into()))?;
        let slot = root
            .children
            .get_mut(index)
            .ok_or_else(|| Error::InvalidFormat("timeline child index is out of range".into()))?;
        *slot = node;
        candidate.stage(persist_id)?;
        *self = candidate;
        Ok(())
    }

    /// Replaces the root timeline and build list atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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

    /// Removes the root timeline child at `index` and returns it.
    ///
    /// # Errors
    ///
    /// Returns an error if the persist ID is unknown, the timeline has no root
    /// node, `index` is out of range, or the rewritten record fails validation.
    pub fn remove(&mut self, persist_id: u32, index: usize) -> Result<ExtendedTimeNode> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        let root = entry
            .extension
            .time_node
            .as_mut()
            .ok_or_else(|| Error::InvalidFormat("timeline has no root node".into()))?;
        if index >= root.children.len() {
            return invalid("timeline child index is out of range");
        }
        let removed = root.children.remove(index);
        candidate.stage(persist_id)?;
        *self = candidate;
        Ok(removed)
    }

    /// Reorders the root timeline children according to `order`.
    ///
    /// # Errors
    ///
    /// Returns an error if the persist ID is unknown, the timeline has no root
    /// node, `order` is not a permutation of every child index, or the
    /// rewritten record fails validation.
    pub fn reorder(&mut self, persist_id: u32, order: &[usize]) -> Result<()> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        let root = entry
            .extension
            .time_node
            .as_mut()
            .ok_or_else(|| Error::InvalidFormat("timeline has no root node".into()))?;
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

    /// Inserts a build-list entry at `index`.
    ///
    /// # Errors
    ///
    /// Returns an error if the persist ID is unknown, `index` is out of range,
    /// or the rewritten record fails validation.
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

    /// Replaces the build-list entry at `index`.
    ///
    /// # Errors
    ///
    /// Returns an error if the persist ID is unknown, the slide/master has no
    /// build list, `index` is out of range, or the rewritten record fails
    /// validation.
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
            .ok_or_else(|| Error::InvalidFormat("slide/master has no build list".into()))?;
        let slot = list
            .builds
            .get_mut(index)
            .ok_or_else(|| Error::InvalidFormat("build index is out of range".into()))?;
        *slot = build;
        candidate.stage(persist_id)?;
        *self = candidate;
        Ok(())
    }

    /// Removes the build-list entry at `index` and returns it.
    ///
    /// # Errors
    ///
    /// Returns an error if the persist ID is unknown, the slide/master has no
    /// build list, `index` is out of range, or the rewritten record fails
    /// validation.
    pub fn remove_build(&mut self, persist_id: u32, index: usize) -> Result<BuildListEntry> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        let list = entry
            .extension
            .build_list
            .as_mut()
            .ok_or_else(|| Error::InvalidFormat("slide/master has no build list".into()))?;
        if index >= list.builds.len() {
            return invalid("build index is out of range");
        }
        let removed = list.builds.remove(index);
        candidate.stage(persist_id)?;
        *self = candidate;
        Ok(removed)
    }

    /// Reorders the build list according to `order`.
    ///
    /// # Errors
    ///
    /// Returns an error if the persist ID is unknown, the slide/master has no
    /// build list, `order` is not a permutation of every entry index, or the
    /// rewritten record fails validation.
    pub fn reorder_builds(&mut self, persist_id: u32, order: &[usize]) -> Result<()> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        let list = entry
            .extension
            .build_list
            .as_mut()
            .ok_or_else(|| Error::InvalidFormat("slide/master has no build list".into()))?;
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

    /// Replaces or clears the legacy `PowerPoint` 97 animation of one shape.
    ///
    /// # Errors
    ///
    /// Returns an error if the persist ID is unknown, the target shape does not
    /// exist, the animation fails serialization, or the shape's `OfficeArt`
    /// container is missing or malformed.
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
            .ok_or_else(|| Error::InvalidFormat("animation persist ID was not found".into()))?;
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
        if let Some(new_animation) = animation {
            let scope = candidate.entries[index].scope;
            candidate.entries[index].legacy.push(LegacyShapeAnimation {
                persist_id,
                scope,
                shape_id,
                animation: new_animation,
            });
            candidate.entries[index]
                .legacy
                .sort_by_key(|value| value.shape_id);
        }
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    /// Finishes the transaction and returns the rewritten package bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the rewritten package cannot be serialized.
    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish()
    }

    fn entry_mut(&mut self, persist_id: u32) -> Result<&mut PersistAnimation> {
        self.entries
            .iter_mut()
            .find(|entry| entry.persist_id == persist_id)
            .ok_or_else(|| Error::InvalidFormat("animation persist ID was not found".into()))
    }

    fn stage(&mut self, persist_id: u32) -> Result<()> {
        let index = self
            .entries
            .iter()
            .position(|entry| entry.persist_id == persist_id)
            .ok_or_else(|| Error::InvalidFormat("animation persist ID was not found".into()))?;
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
        let (parsed, consumed) = Record::parse(&record, 0)?;
        if consumed != record.len() || parsed.data_length as usize + 8 != record.len() {
            return corrupted("rewritten slide/master record failed length validation");
        }
        let reparsed_payload = find_ppt10_payload(&parsed)?
            .ok_or_else(|| Error::Corrupted("rewritten ___PPT10 payload is missing".into()))?;
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

fn find_ppt10_payload(record: &Record) -> Result<Option<Vec<u8>>> {
    for prog_tags in record.find_children(RecordType::ProgTags) {
        for binary in prog_tags.find_children(RecordType::ProgBinaryTag) {
            let Some(name) = binary.find_child(RecordType::CString) else {
                continue;
            };
            if is_ppt10_name(&name.data) {
                let data = binary
                    .find_child(RecordType::BinaryTagData)
                    .ok_or_else(|| {
                        Error::Corrupted("___PPT10 tag is missing BinaryTagData".into())
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

pub(super) fn rewrite_extension_payload(
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
        if kind == RecordType::ExtTimeNode.as_u16() {
            if saw_root {
                return invalid("extension contains duplicate root timelines");
            }
            saw_root = true;
            if let Some(value) = root {
                output.extend(write_extended_time_node(value)?);
            }
        } else if kind == RecordType::BuildList.as_u16() {
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
        if raw_type(record).ok() != Some(RecordType::ProgBinaryTag.as_u16()) {
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
        RecordType::CString.as_u16(),
        &"___PPT10"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>(),
    )?;
    binary_children.extend(atom(0, 0, RecordType::BinaryTagData.as_u16(), payload)?);
    let binary = container(0, RecordType::ProgBinaryTag.as_u16(), &binary_children)?;
    let tags = container(0, RecordType::ProgTags.as_u16(), &binary)?;
    append_container_child(&rewritten, &tags)
}

fn prog_binary_is_ppt10(record: &[u8]) -> Result<bool> {
    let mut offset = 8;
    while offset < record.len() {
        let child = raw_record(record, offset)?;
        if raw_type(child)? == RecordType::CString.as_u16() {
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
        if raw_type(child)? == RecordType::BinaryTagData.as_u16() {
            if found {
                return invalid("___PPT10 contains duplicate BinaryTagData");
            }
            out.extend(atom(0, 0, RecordType::BinaryTagData.as_u16(), payload)?);
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

pub(super) fn collect_shapes_and_legacy(
    persist_id: u32,
    scope: Scope,
    record: &Record,
    limits: EditorLimits,
) -> Result<(BTreeSet<u32>, Vec<LegacyShapeAnimation>)> {
    let mut ids = BTreeSet::new();
    let mut legacy = Vec::new();
    for drawing in record.find_children(RecordType::PPDrawing) {
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
                            if raw_type(item)? == RecordType::AnimationInfo.as_u16() {
                                let (parsed, used) = Record::parse(item, 0)?;
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

pub(super) fn rewrite_shape_animation(
    root: &[u8],
    shape_id: u32,
    animation: Option<&[u8]>,
) -> Result<(Vec<u8>, bool)> {
    rewrite_ppt_record(root, &mut |record| {
        if raw_type(record).ok() != Some(RecordType::PPDrawing.as_u16()) {
            return Ok(None);
        }
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
                if raw_type(item)? == RecordType::AnimationInfo.as_u16() {
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
        .ok_or_else(|| Error::Corrupted("container header is truncated".into()))?
        .to_vec();
    payload.extend_from_slice(child);
    rebuild_record(record, &payload)
}
fn rebuild_record(record: &[u8], payload: &[u8]) -> Result<Vec<u8>> {
    let mut out = record
        .get(..4)
        .ok_or_else(|| Error::Corrupted("record header is truncated".into()))?
        .to_vec();
    out.extend_from_slice(
        &u32::try_from(payload.len())
            .map_err(|_err| Error::InvalidFormat("record payload exceeds u32".into()))?
            .to_le_bytes(),
    );
    out.extend_from_slice(payload);
    Ok(out)
}
pub(super) fn atom(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Result<Vec<u8>> {
    record_bytes(version, instance, kind, payload)
}
pub(super) fn container(instance: u16, kind: u16, payload: &[u8]) -> Result<Vec<u8>> {
    record_bytes(0x0F, instance, kind, payload)
}
pub(super) fn escher_record(
    version: u16,
    instance: u16,
    kind: u16,
    payload: &[u8],
) -> Result<Vec<u8>> {
    record_bytes(version, instance, kind, payload)
}
fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Result<Vec<u8>> {
    let mut out = ((instance << 4) | (version & 0xF)).to_le_bytes().to_vec();
    out.extend_from_slice(&kind.to_le_bytes());
    out.extend_from_slice(
        &u32::try_from(payload.len())
            .map_err(|_err| Error::InvalidFormat("record payload exceeds u32".into()))?
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
        .ok_or_else(|| Error::Corrupted("record length overflow".into()))?;
    data.get(offset..end)
        .ok_or_else(|| Error::Corrupted("record is truncated".into()))
}
fn raw_type(record: &[u8]) -> Result<u16> {
    record
        .get(2..4)
        .map(|v| u16::from_le_bytes([v[0], v[1]]))
        .ok_or_else(|| Error::Corrupted("record type is truncated".into()))
}
fn raw_version(record: &[u8]) -> Result<u16> {
    record
        .get(0..2)
        .map(|v| u16::from_le_bytes([v[0], v[1]]) & 0xF)
        .ok_or_else(|| Error::Corrupted("record version is truncated".into()))
}
fn u32_at(data: &[u8], offset: usize) -> Result<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| Error::Corrupted("record u32 is truncated".into()))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
