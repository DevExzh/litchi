use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, TargetMode};

use super::super::model::{Frame, Kind, Mode, Object, Target};
use super::codec::{self, Edit, Located};
use super::model::Definition;
use super::validation::{
    MAX_OBJECTS, MAX_RELATIONSHIPS, validate_anchor, validate_kind, validate_name,
    validate_payload, validate_program, validate_source, validate_target, validate_text,
};
use crate::{Error, Result};

/// Stable fingerprint of a complete slide-owned OLE source graph.
pub type Revision = u64;

/// Relationship state retained for source checks and exact graph publication.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RelationshipState {
    pub(crate) id: String,
    pub(crate) relationship_type: String,
    pub(crate) target_ref: String,
    pub(crate) target_mode: TargetMode,
}

/// Opaque OPC part state referenced by the selected slide's OLE graph.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct PartSource {
    pub(crate) part_name: PackURI,
    pub(crate) content_type: String,
    pub(crate) bytes: Arc<Vec<u8>>,
    pub(crate) relationships: Arc<Vec<RelationshipState>>,
}

/// Immutable, detached view of all OLE objects owned by one slide.
#[derive(Debug, Clone)]
pub struct Snapshot {
    pub(crate) slide_index: usize,
    pub(crate) slide_part_name: PackURI,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) relationships: Arc<Vec<RelationshipState>>,
    pub(crate) objects: Arc<Vec<Object>>,
    pub(crate) parts: Arc<Vec<PartSource>>,
    pub(crate) package_part_names: Arc<Vec<PackURI>>,
    pub(crate) revision: Revision,
}

impl Snapshot {
    /// Load one slide's complete inert OLE graph by absolute slide part name.
    pub fn load(
        package: &OpcPackage,
        slide_index: usize,
        slide_part_name: &PackURI,
    ) -> Result<Self> {
        super::package::load(package, slide_index, slide_part_name)
    }

    pub(crate) fn from_parts(
        slide_index: usize,
        slide_part_name: PackURI,
        source_xml: Arc<Vec<u8>>,
        relationships: Vec<RelationshipState>,
        objects: Vec<Object>,
        parts: Vec<PartSource>,
        package_part_names: Vec<PackURI>,
    ) -> Result<Self> {
        validate_source(source_xml.as_slice())?;
        if objects.len() > MAX_OBJECTS {
            return Err(crate::presentation::embedded::limit(
                "OLE object count",
                MAX_OBJECTS,
            ));
        }
        let relationships = Arc::new(sorted_relationships(relationships)?);
        let parts = Arc::new(sorted_parts(parts)?);
        let mut package_part_names = package_part_names;
        package_part_names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
        package_part_names.dedup();
        let package_part_names = Arc::new(package_part_names);
        let objects = Arc::new(objects);
        let revision = fingerprint(
            source_xml.as_slice(),
            relationships.as_ref(),
            objects.as_ref(),
            parts.as_ref(),
        );
        Ok(Self {
            slide_index,
            slide_part_name,
            source_xml,
            relationships,
            objects,
            parts,
            package_part_names,
            revision,
        })
    }

    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    pub fn slide_part_name(&self) -> &PackURI {
        &self.slide_part_name
    }

    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    pub fn objects(&self) -> &[Object] {
        self.objects.as_slice()
    }

    pub fn revision(&self) -> Revision {
        self.revision
    }

    /// Return an exact inert payload for an embedded object, if it has one.
    pub fn payload(&self, index: usize) -> Result<Option<&[u8]>> {
        let object = self
            .objects
            .get(index)
            .ok_or_else(|| index_error(index, self.objects.len()))?;
        let Some(Target::Internal { part_name, .. }) = object.target.as_ref() else {
            return Ok(None);
        };
        Ok(self
            .parts
            .iter()
            .find(|part| part.part_name == *part_name)
            .map(|part| part.bytes.as_slice()))
    }

    pub fn edit(&self) -> Transaction {
        let parts: HashMap<PackURI, &PartSource> = self
            .parts
            .iter()
            .map(|part| (part.part_name.clone(), part))
            .collect();
        let working = self
            .objects
            .iter()
            .enumerate()
            .map(|(source_index, object)| Working {
                source_index: Some(source_index),
                object: object.clone(),
                payload: object.target.as_ref().and_then(|target| match target {
                    Target::Internal { part_name, .. } => {
                        parts.get(part_name).map(|part| Arc::clone(&part.bytes))
                    },
                    Target::External { .. } => None,
                }),
            })
            .collect();
        Transaction {
            source: self.clone(),
            working,
        }
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.slide_index == other.slide_index
            && self.slide_part_name == other.slide_part_name
            && self.source_xml.as_slice() == other.source_xml.as_slice()
            && self.relationships == other.relationships
            && self.objects == other.objects
            && self.parts == other.parts
    }
}

#[derive(Debug, Clone)]
struct Working {
    source_index: Option<usize>,
    object: Object,
    payload: Option<Arc<Vec<u8>>>,
}

/// Failure-atomic editor over one slide-owned OLE object collection.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    working: Vec<Working>,
}

impl Transaction {
    pub fn source(&self) -> &Snapshot {
        &self.source
    }

    pub fn objects(&self) -> impl Iterator<Item = &Object> {
        self.working.iter().map(|value| &value.object)
    }

    pub fn object(&self, index: usize) -> Result<&Object> {
        self.working
            .get(index)
            .map(|value| &value.object)
            .ok_or_else(|| index_error(index, self.working.len()))
    }

    pub fn is_changed(&self) -> bool {
        if self.working.len() != self.source.objects.len() {
            return true;
        }
        self.working.iter().enumerate().any(|(index, value)| {
            value.source_index != Some(index)
                || value.object != self.source.objects[index]
                || value.payload.as_ref().map(|value| value.as_slice())
                    != self.source.payload(index).ok().flatten()
        })
    }

    /// Add an inert embedded or linked OLE object and return its transaction index.
    pub fn add(&mut self, definition: Definition) -> Result<usize> {
        validate_definition(&definition)?;
        if self.working.len() >= MAX_OBJECTS {
            return Err(crate::presentation::embedded::limit(
                "OLE object count",
                MAX_OBJECTS,
            ));
        }
        let object = Object {
            slide_index: self.source.slide_index,
            index: self.working.len(),
            shape_id: None,
            shape_name: None,
            legacy_shape_id: None,
            name: definition.name.clone(),
            program_id: definition.program_id.clone(),
            show_as_icon: definition.show_as_icon,
            preview_width: definition.preview_width,
            preview_height: definition.preview_height,
            anchor: Some(definition.anchor),
            mode: definition.mode,
            relationship_id: None,
            kind: Some(definition.kind),
            target: definition.target.clone().map(|target| Target::External {
                target,
                relationship_type: relationship_type(definition.kind).to_owned(),
            }),
            preview_relationship_id: None,
        };
        let payload = definition.payload.map(Arc::new);
        self.working.push(Working {
            source_index: None,
            object,
            payload,
        });
        Ok(self.working.len() - 1)
    }

    /// Remove the selected graphic frame and collect any now-unreferenced payload.
    pub fn remove(&mut self, index: usize) -> Result<Object> {
        let value = self
            .working
            .get(index)
            .cloned()
            .ok_or_else(|| index_error(index, self.working.len()))?;
        self.working.remove(index);
        self.normalize_indices();
        Ok(value.object)
    }

    /// Detach the selected OLE graph from the slide.  This is intentionally
    /// distinct in the API from metadata edits, while publication has the same
    /// safe graph semantics as removal: no orphan payload is retained.
    pub fn detach(&mut self, index: usize) -> Result<Object> {
        self.remove(index)
    }

    /// Replace all typed metadata and the opaque payload/target in place.
    pub fn replace(&mut self, index: usize, definition: Definition) -> Result<()> {
        validate_definition(&definition)?;
        let value = self.working_mut(index)?;
        if value.object.kind != Some(definition.kind) || value.object.mode != definition.mode {
            return Err(invalid(
                "OLE replacement cannot change kind or embedded/link mode",
            ));
        }
        value.object.anchor = Some(definition.anchor);
        value.object.name = definition.name;
        value.object.program_id = definition.program_id;
        value.object.show_as_icon = definition.show_as_icon;
        value.object.preview_width = definition.preview_width;
        value.object.preview_height = definition.preview_height;
        if definition.mode == Mode::Linked {
            let target = definition.target.expect("validated linked target");
            let relationship_type = relationship_type(definition.kind).to_owned();
            value.object.target = Some(Target::External {
                target,
                relationship_type,
            });
        }
        value.payload = definition.payload.map(Arc::new);
        Ok(())
    }

    pub fn set_anchor(&mut self, index: usize, anchor: Frame) -> Result<bool> {
        validate_anchor(anchor)?;
        let value = self.working_mut(index)?;
        if value.object.anchor == Some(anchor) {
            return Ok(false);
        }
        value.object.anchor = Some(anchor);
        Ok(true)
    }

    pub fn set_name(&mut self, index: usize, name: Option<String>) -> Result<bool> {
        if let Some(value) = name.as_deref() {
            validate_name(value)?;
        }
        let value = self.working_mut(index)?;
        if value.object.name == name {
            return Ok(false);
        }
        value.object.name = name;
        Ok(true)
    }

    pub fn set_program_id(&mut self, index: usize, program_id: Option<String>) -> Result<bool> {
        if let Some(value) = program_id.as_deref() {
            validate_program(value)?;
        }
        let value = self.working_mut(index)?;
        if value.object.program_id == program_id {
            return Ok(false);
        }
        value.object.program_id = program_id;
        Ok(true)
    }

    pub fn set_show_as_icon(&mut self, index: usize, value: Option<bool>) -> Result<bool> {
        let object = self.working_mut(index)?;
        if object.object.show_as_icon == value {
            return Ok(false);
        }
        object.object.show_as_icon = value;
        Ok(true)
    }

    pub fn set_preview_size(&mut self, index: usize, value: Option<(u32, u32)>) -> Result<bool> {
        if let Some((width, height)) = value
            && (width == 0 || height == 0)
        {
            return Err(invalid("OLE preview dimensions must be positive"));
        }
        let object = self.working_mut(index)?;
        let before = (object.object.preview_width, object.object.preview_height);
        object.object.preview_width = value.map(|value| value.0);
        object.object.preview_height = value.map(|value| value.1);
        Ok(before != (object.object.preview_width, object.object.preview_height))
    }

    /// Replace an embedded payload without interpreting its OLE bytes.
    pub fn replace_payload(&mut self, index: usize, payload: impl Into<Vec<u8>>) -> Result<()> {
        let payload = payload.into();
        validate_payload(&payload)?;
        let object = self.working_mut(index)?;
        if object.object.mode != Mode::Embedded {
            return Err(invalid("linked OLE objects do not have embedded payloads"));
        }
        object.payload = Some(Arc::new(payload));
        Ok(())
    }

    /// Change a linked target while retaining it as an inert external URI.
    pub fn set_link_target(&mut self, index: usize, target: impl Into<String>) -> Result<()> {
        let target = target.into();
        validate_text(&target, "OLE link target", false)?;
        let object = self.working_mut(index)?;
        if object.object.mode != Mode::Linked {
            return Err(invalid(
                "embedded OLE objects do not have external link targets",
            ));
        }
        let relationship_type = object
            .object
            .kind
            .map(relationship_type)
            .ok_or_else(|| invalid("OLE object kind is missing"))?
            .to_owned();
        object.object.target = Some(Target::External {
            target,
            relationship_type,
        });
        Ok(())
    }

    /// Validate and consume this edit into a reversible package patch.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch {
                before: self.source.clone(),
                after: self.source.clone(),
            };
            return Ok(Commit {
                snapshot: self.source,
                patch,
            });
        }
        let build = build_after(&self.source, &self.working)?;
        let snapshot = Snapshot::from_parts(
            self.source.slide_index,
            self.source.slide_part_name.clone(),
            Arc::new(build.source_xml),
            build.relationships,
            build.objects,
            build.parts,
            self.source.package_part_names.as_ref().clone(),
        )?;
        let patch = Patch {
            before: self.source,
            after: snapshot.clone(),
        };
        Ok(Commit { snapshot, patch })
    }

    pub fn rollback(self) -> Snapshot {
        self.source
    }

    fn normalize_indices(&mut self) {
        for (index, value) in self.working.iter_mut().enumerate() {
            value.object.index = index;
            value.object.slide_index = self.source.slide_index;
        }
    }

    fn working_mut(&mut self, index: usize) -> Result<&mut Working> {
        let len = self.working.len();
        self.working
            .get_mut(index)
            .ok_or_else(|| index_error(index, len))
    }
}

/// A successful detached OLE edit and its reversible source-checked patch.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    pub fn is_changed(&self) -> bool {
        !self.patch.is_empty()
    }

    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// Reversible, source-checked replacement of one slide-owned OLE graph.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    pub fn is_changed(&self) -> bool {
        !self.is_empty()
    }

    pub fn expected_revision(&self) -> Revision {
        self.before.revision
    }

    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    pub fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        super::package::apply_patch(package, self)
    }

    pub fn undo(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        self.inverse().apply(package)
    }
}

struct Build {
    source_xml: Vec<u8>,
    relationships: Vec<RelationshipState>,
    objects: Vec<Object>,
    parts: Vec<PartSource>,
}

fn build_after(source: &Snapshot, working: &[Working]) -> Result<Build> {
    if working.len() > MAX_OBJECTS {
        return Err(crate::presentation::embedded::limit(
            "OLE object count",
            MAX_OBJECTS,
        ));
    }
    let document = codec::locate(source.source_xml.as_slice())?;
    if document.frames.len() != source.objects.len() {
        return Err(invalid("OLE source object order changed during capture"));
    }
    let mut used_relationships: HashSet<String> = source
        .relationships
        .iter()
        .map(|value| value.id.clone())
        .collect();
    let mut used_parts: HashSet<PackURI> = source.package_part_names.iter().cloned().collect();
    let mut next_shape_id = document.max_shape_id;
    let mut relationships = source.relationships.as_ref().clone();
    let source_part_map: HashMap<PackURI, PartSource> = source
        .parts
        .iter()
        .cloned()
        .map(|part| (part.part_name.clone(), part))
        .collect();
    let mut parts: HashMap<PackURI, PartSource> = HashMap::new();
    let mut resolved = Vec::with_capacity(working.len());
    let mut edits = Vec::new();
    let mut additions = Vec::new();

    for (index, value) in working.iter().enumerate() {
        let source_object = value
            .source_index
            .and_then(|source_index| source.objects.get(source_index));
        let mut object = value.object.clone();
        let kind = object
            .kind
            .ok_or_else(|| invalid("OLE object kind is missing"))?;
        validate_kind(kind)?;
        let payload = value.payload.as_ref().map(|value| value.as_slice());
        validate_target(
            object.mode,
            match object.target.as_ref() {
                Some(Target::External { target, .. }) => Some(target.as_str()),
                _ => None,
            },
            payload,
        )?;
        if let Some(name) = object.name.as_deref() {
            validate_name(name)?;
        }
        if let Some(program_id) = object.program_id.as_deref() {
            validate_program(program_id)?;
        }
        validate_anchor(
            object
                .anchor
                .ok_or_else(|| invalid("OLE anchor is missing"))?,
        )?;

        let relationship_id = if let Some(id) = object.relationship_id.clone() {
            id
        } else {
            let id = next_relationship_id(&mut used_relationships);
            object.relationship_id = Some(id.clone());
            id
        };
        let relationship_type = relationship_type(kind).to_owned();
        let (target, target_ref, target_mode, part_name) = match object.mode {
            Mode::Embedded => {
                let part_name = match object.target.as_ref() {
                    Some(Target::Internal { part_name, .. }) => part_name.clone(),
                    Some(Target::External { .. }) => {
                        return Err(invalid("embedded OLE object has an external target"));
                    },
                    None => next_part_name(&mut used_parts, kind)?,
                };
                let content_type = content_type(kind).to_owned();
                let target_ref = part_name.relative_ref(source.slide_part_name.base_uri());
                let target = Target::Internal {
                    part_name: part_name.clone(),
                    content_type: content_type.clone(),
                    relationship_type: relationship_type.clone(),
                };
                let bytes = value
                    .payload
                    .clone()
                    .ok_or_else(|| invalid("embedded OLE object has no payload"))?;
                let relationships = source_part_map
                    .get(&part_name)
                    .map(|part| Arc::clone(&part.relationships))
                    .unwrap_or_else(|| Arc::new(Vec::new()));
                parts.insert(
                    part_name.clone(),
                    PartSource {
                        part_name: part_name.clone(),
                        content_type,
                        bytes,
                        relationships,
                    },
                );
                (target, target_ref, TargetMode::Internal, Some(part_name))
            },
            Mode::Linked => {
                let target = match object.target.as_ref() {
                    Some(Target::External { target, .. }) => target.clone(),
                    Some(Target::Internal { .. }) => {
                        return Err(invalid(
                            "linked OLE replacement requires an external target",
                        ));
                    },
                    None => return Err(invalid("linked OLE object has no target")),
                };
                (
                    Target::External {
                        target: target.clone(),
                        relationship_type: relationship_type.clone(),
                    },
                    target,
                    TargetMode::External,
                    None,
                )
            },
        };
        object.kind = Some(kind);
        object.target = Some(target.clone());
        object.index = index;
        object.slide_index = source.slide_index;

        if let Some(source_object) = source_object {
            let source_index = value.source_index.expect("source object index");
            let located = document
                .frames
                .get(source_index)
                .ok_or_else(|| invalid("OLE source frame disappeared"))?;
            if source_object != &object {
                add_object_edits(
                    &source.source_xml,
                    located,
                    source_object,
                    &object,
                    &mut edits,
                )?;
            }
            if source_object.relationship_id.as_deref() != Some(relationship_id.as_str()) {
                return Err(invalid("OLE relationship identity cannot change"));
            }
        } else {
            next_shape_id = next_shape_id
                .checked_add(1)
                .ok_or_else(|| invalid("OLE shape ID overflow"))?;
            object.shape_id = Some(next_shape_id);
            object.shape_name = Some(
                object
                    .name
                    .clone()
                    .unwrap_or_else(|| format!("OLE Object {next_shape_id}")),
            );
            additions.push(frame_xml(&object, &relationship_id, next_shape_id));
        }
        upsert_relationship(
            &mut relationships,
            &relationship_id,
            relationship_type,
            target_ref,
            target_mode,
        )?;
        if let Some(part_name) = part_name {
            if let Some(part) = parts.get_mut(&part_name) {
                part.bytes = value
                    .payload
                    .clone()
                    .ok_or_else(|| invalid("embedded OLE object has no payload"))?;
            }
        }
        resolved.push(object);
    }

    let kept_ids: HashSet<String> = resolved
        .iter()
        .filter_map(|object| object.relationship_id.clone())
        .chain(
            resolved
                .iter()
                .filter_map(|object| object.preview_relationship_id.clone()),
        )
        .collect();
    relationships.retain(|relationship| {
        let was_preview = source.objects.iter().any(|object| {
            object.preview_relationship_id.as_deref() == Some(relationship.id.as_str())
        });
        (!is_ole_relationship(&relationship.relationship_type) && !was_preview)
            || kept_ids.contains(&relationship.id)
    });
    for object in &resolved {
        if let Some(Target::Internal { part_name, .. }) = object.target.as_ref()
            && let Some(part) = source_part_map.get(part_name)
            && !parts.contains_key(part_name)
        {
            parts.insert(part_name.clone(), part.clone());
        }
        if let Some(id) = object.preview_relationship_id.as_deref()
            && let Some(relationship) = source.relationships.iter().find(|value| value.id == id)
            && relationship.target_mode == TargetMode::Internal
            && let Ok(part_name) = PackURI::from_rel_ref(
                source.slide_part_name.base_uri(),
                relationship
                    .target_ref
                    .split(['?', '#'])
                    .next()
                    .unwrap_or_default(),
            )
            && let Some(part) = source_part_map.get(&part_name)
            && !parts.contains_key(&part_name)
        {
            parts.insert(part_name, part.clone());
        }
    }

    if !additions.is_empty() {
        let fragment = additions.join("");
        edits.push(codec::append_fragment(
            source.source_xml.as_slice(),
            document.insertion,
            fragment.as_bytes(),
        )?);
    }
    let kept: HashSet<usize> = source
        .objects
        .iter()
        .enumerate()
        .filter_map(|(index, _)| {
            working
                .iter()
                .any(|value| value.source_index == Some(index))
                .then_some(index)
        })
        .collect();
    for (index, _) in source.objects.iter().enumerate() {
        if !kept.contains(&index) {
            edits.push(Edit {
                range: document.frames[index].frame.clone(),
                replacement: Vec::new(),
            });
        }
    }
    let source_xml = codec::apply_edits(source.source_xml.as_slice(), edits)?;
    let _ = codec::locate(&source_xml)?;
    Ok(Build {
        source_xml,
        relationships,
        objects: resolved,
        parts: parts.into_values().collect(),
    })
}

fn add_object_edits(
    source: &[u8],
    located: &Located,
    before: &Object,
    after: &Object,
    edits: &mut Vec<Edit>,
) -> Result<()> {
    if before.name != after.name {
        if let Some(edit) = codec::attribute_edit(
            source,
            &located.object,
            b"name",
            false,
            after.name.as_deref(),
        )? {
            edits.push(edit);
        }
    }
    if before.program_id != after.program_id {
        if let Some(edit) = codec::attribute_edit(
            source,
            &located.object,
            b"progId",
            false,
            after.program_id.as_deref(),
        )? {
            edits.push(edit);
        }
    }
    if before.show_as_icon != after.show_as_icon {
        let value = after
            .show_as_icon
            .map(|value| if value { "true" } else { "false" });
        if let Some(edit) =
            codec::attribute_edit(source, &located.object, b"showAsIcon", false, value)?
        {
            edits.push(edit);
        }
    }
    if before.preview_width != after.preview_width {
        let value = after.preview_width.map(|value| value.to_string());
        if let Some(edit) =
            codec::attribute_edit(source, &located.object, b"imgW", false, value.as_deref())?
        {
            edits.push(edit);
        }
    }
    if before.preview_height != after.preview_height {
        let value = after.preview_height.map(|value| value.to_string());
        if let Some(edit) =
            codec::attribute_edit(source, &located.object, b"imgH", false, value.as_deref())?
        {
            edits.push(edit);
        }
    }
    if before.anchor != after.anchor {
        let (Some((off, ext)), Some(anchor)) = (located.anchor.as_ref(), after.anchor) else {
            return Err(invalid("OLE anchor source is incomplete"));
        };
        for (node, local, value) in [
            (off, b"x".as_slice(), anchor.x.to_string()),
            (off, b"y".as_slice(), anchor.y.to_string()),
            (ext, b"cx".as_slice(), anchor.cx.to_string()),
            (ext, b"cy".as_slice(), anchor.cy.to_string()),
        ] {
            if let Some(edit) = codec::attribute_edit(source, node, local, false, Some(&value))? {
                edits.push(edit);
            }
        }
    }
    Ok(())
}

fn validate_definition(value: &Definition) -> Result<()> {
    validate_kind(value.kind)?;
    validate_anchor(value.anchor)?;
    if let Some(name) = value.name.as_deref() {
        validate_name(name)?;
    }
    if let Some(program_id) = value.program_id.as_deref() {
        validate_program(program_id)?;
    }
    validate_target(
        value.mode,
        value.target.as_deref(),
        value.payload.as_deref(),
    )
}

fn relationship_type(kind: Kind) -> &'static str {
    match kind {
        Kind::OleObject => litchi_opc::constants::relationship_type::OLE_OBJECT,
        Kind::Package => litchi_opc::constants::relationship_type::PACKAGE,
    }
}

fn content_type(kind: Kind) -> &'static str {
    match kind {
        Kind::OleObject => litchi_opc::constants::content_type::OFC_OLE_OBJECT,
        Kind::Package => litchi_opc::constants::content_type::OFC_PACKAGE,
    }
}

fn is_ole_relationship(value: &str) -> bool {
    matches!(
        value,
        litchi_opc::constants::relationship_type::OLE_OBJECT
            | litchi_opc::constants::relationship_type::PACKAGE
            | litchi_opc::constants::relationship_type::STRICT_OLE_OBJECT
            | litchi_opc::constants::relationship_type::STRICT_PACKAGE
    )
}

fn upsert_relationship(
    relationships: &mut Vec<RelationshipState>,
    id: &str,
    relationship_type: String,
    target_ref: String,
    target_mode: TargetMode,
) -> Result<()> {
    if let Some(current) = relationships.iter_mut().find(|value| value.id == id) {
        current.relationship_type = relationship_type;
        current.target_ref = target_ref;
        current.target_mode = target_mode;
        return Ok(());
    }
    if relationships.len() >= MAX_RELATIONSHIPS {
        return Err(crate::presentation::embedded::limit(
            "OLE relationship count",
            MAX_RELATIONSHIPS,
        ));
    }
    relationships.push(RelationshipState {
        id: id.to_owned(),
        relationship_type,
        target_ref,
        target_mode,
    });
    Ok(())
}

fn next_relationship_id(used: &mut HashSet<String>) -> String {
    for index in 1..=u32::MAX {
        let value = format!("rId{index}");
        if used.insert(value.clone()) {
            return value;
        }
    }
    unreachable!("relationship ID namespace exhausted")
}

fn next_part_name(used: &mut HashSet<PackURI>, kind: Kind) -> Result<PackURI> {
    let template = match kind {
        Kind::OleObject => "/ppt/embeddings/oleObject{}.bin",
        Kind::Package => "/ppt/embeddings/package{}.bin",
    };
    for index in 1..1_000_000u32 {
        let name = PackURI::new(template.replace("{}", &index.to_string())).map_err(Error::Uri)?;
        if used.insert(name.clone()) {
            return Ok(name);
        }
    }
    Err(crate::presentation::embedded::limit(
        "OLE embedding part namespace",
        1_000_000,
    ))
}

fn frame_xml(object: &Object, relationship_id: &str, shape_id: u32) -> String {
    let anchor = object.anchor.expect("validated OLE anchor");
    let name = object
        .name
        .clone()
        .unwrap_or_else(|| format!("OLE Object {shape_id}"));
    let mut output = format!(
        r#"<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="{shape_id}" name="{}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></p:xfrm><a:graphic><a:graphicData uri="{}"><p:oleObj"#,
        escape(&name),
        anchor.x,
        anchor.y,
        anchor.cx,
        anchor.cy,
        super::super::codec::OLE_GRAPHIC_DATA_URI,
    );
    if let Some(name) = object.name.as_deref() {
        output.push_str(&format!(r#" name="{}""#, escape(name)));
    }
    if let Some(program_id) = object.program_id.as_deref() {
        output.push_str(&format!(r#" progId="{}""#, escape(program_id)));
    }
    if let Some(value) = object.show_as_icon {
        output.push_str(&format!(
            r#" showAsIcon="{}""#,
            if value { "true" } else { "false" }
        ));
    }
    if let Some(value) = object.preview_width {
        output.push_str(&format!(r#" imgW="{value}""#));
    }
    if let Some(value) = object.preview_height {
        output.push_str(&format!(r#" imgH="{value}""#));
    }
    output.push_str(&format!(r#" r:id="{}">"#, escape(relationship_id)));
    output.push_str(if object.mode == Mode::Embedded {
        "<p:embed/>"
    } else {
        "<p:link/>"
    });
    output.push_str("</p:oleObj></a:graphicData></a:graphic></p:graphicFrame>");
    output
}

pub(crate) fn sorted_relationships(
    mut values: Vec<RelationshipState>,
) -> Result<Vec<RelationshipState>> {
    if values.len() > MAX_RELATIONSHIPS {
        return Err(crate::presentation::embedded::limit(
            "OLE relationship count",
            MAX_RELATIONSHIPS,
        ));
    }
    values.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    for pair in values.windows(2) {
        if pair[0].id == pair[1].id {
            return Err(invalid("duplicate OLE relationship ID"));
        }
    }
    Ok(values)
}

pub(crate) fn sorted_parts(mut values: Vec<PartSource>) -> Result<Vec<PartSource>> {
    values.sort_unstable_by(|left, right| left.part_name.as_str().cmp(right.part_name.as_str()));
    for part in &values {
        if part.bytes.len() > super::validation::MAX_PART_BYTES {
            return Err(crate::presentation::embedded::limit(
                "OLE payload bytes",
                super::validation::MAX_PART_BYTES,
            ));
        }
    }
    for pair in values.windows(2) {
        if pair[0].part_name == pair[1].part_name {
            return Err(invalid("duplicate OLE payload part"));
        }
    }
    Ok(values)
}

fn fingerprint(
    source: &[u8],
    relationships: &[RelationshipState],
    objects: &[Object],
    parts: &[PartSource],
) -> Revision {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    feed_bytes(&mut hash, source);
    for relationship in relationships {
        feed_text(&mut hash, &relationship.id);
        feed_text(&mut hash, &relationship.relationship_type);
        feed_text(&mut hash, &relationship.target_ref);
        hash ^= match relationship.target_mode {
            TargetMode::Internal => 1,
            TargetMode::External => 2,
        };
    }
    for object in objects {
        hash ^= object.index as u64;
        if let Some(id) = object.shape_id {
            hash ^= u64::from(id);
        }
        if let Some(id) = object.relationship_id.as_deref() {
            feed_text(&mut hash, id);
        }
        if let Some(name) = object.name.as_deref() {
            feed_text(&mut hash, name);
        }
    }
    for part in parts {
        feed_text(&mut hash, part.part_name.as_str());
        feed_text(&mut hash, &part.content_type);
        feed_bytes(&mut hash, part.bytes.as_slice());
        for relationship in part.relationships.iter() {
            feed_text(&mut hash, &relationship.id);
            feed_text(&mut hash, &relationship.target_ref);
        }
    }
    hash
}

fn feed_bytes(hash: &mut u64, value: &[u8]) {
    for byte in value {
        *hash ^= u64::from(*byte);
        *hash = hash.wrapping_mul(0x100000001b3);
    }
}

fn feed_text(hash: &mut u64, value: &str) {
    feed_bytes(hash, value.as_bytes());
    *hash ^= 0xff;
}

fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn index_error(index: usize, len: usize) -> Error {
    Error::Invalid(format!(
        "OLE object index {index} is outside a collection of length {len}"
    ))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
