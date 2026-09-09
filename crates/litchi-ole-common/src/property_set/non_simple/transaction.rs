//! Failure-atomic edits and source-checked patches for non-simple storages.

use super::codec::{self, PayloadEdit};
use super::model::{ElementKind, EntryRecord, Limits, Snapshot, State, same_path};
use crate::object::directory::{EntryKind, Metadata};
use crate::property_set::model::{
    IndirectPropertyName, Stream, Value, VersionedStream, invalid, try_clone_property_set,
};
use litchi_cfb::{OleError, directory_names_equal, validate_directory_name};
use std::sync::Arc;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    fn of(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }

    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }

    #[must_use]
    pub const fn fingerprint(self) -> u64 {
        self.value()
    }
}

#[derive(Debug, Clone)]
pub struct Editor {
    source: Snapshot,
    candidate: Option<Stream>,
    records: Option<Vec<EntryRecord>>,
    edits: Vec<PayloadEdit>,
}

impl Editor {
    pub(crate) fn new(source: Snapshot) -> Self {
        Self {
            source,
            candidate: None,
            records: None,
            edits: Vec::new(),
        }
    }

    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.candidate.is_some() || self.records.is_some() || !self.edits.is_empty()
    }

    pub fn contents(&self) -> Result<Stream, OleError> {
        match &self.candidate {
            Some(candidate) => clone_stream(candidate),
            None => self.source.contents(),
        }
    }

    pub fn update<F>(&mut self, edit: F) -> Result<&mut Self, OleError>
    where
        F: FnOnce(&mut Stream) -> Result<(), OleError>,
    {
        let mut candidate = self.ensure_contents()?;
        edit(&mut candidate)?;
        self.set_candidate(candidate)?;
        Ok(self)
    }

    pub fn replace_contents(&mut self, candidate: Stream) -> Result<&mut Self, OleError> {
        self.set_candidate(candidate)?;
        Ok(self)
    }

    pub fn set_property(
        &mut self,
        format_identifier: crate::property_set::Guid,
        property_identifier: u32,
        value: Value,
    ) -> Result<&mut Self, OleError> {
        self.update(|stream| {
            let section = stream
                .section_mut(format_identifier)
                .ok_or_else(|| invalid("Property Set section is not present"))?;
            if section.property(property_identifier).is_some() {
                section.update(property_identifier, value).map(|_| ())
            } else {
                section.add(property_identifier, value)
            }
        })
    }

    pub fn remove_property(
        &mut self,
        format_identifier: crate::property_set::Guid,
        property_identifier: u32,
    ) -> Result<Option<Value>, OleError> {
        let mut candidate = self.ensure_contents()?;
        let removed = candidate
            .section_mut(format_identifier)
            .ok_or_else(|| invalid("Property Set section is not present"))?
            .remove(property_identifier);
        self.set_candidate(candidate)?;
        Ok(removed)
    }

    pub fn set_stream_property(
        &mut self,
        format_identifier: crate::property_set::Guid,
        property_identifier: u32,
    ) -> Result<&mut Self, OleError> {
        self.set_property(
            format_identifier,
            property_identifier,
            Value::Stream(IndirectPropertyName::new(property_identifier)?),
        )
    }

    pub fn set_storage_property(
        &mut self,
        format_identifier: crate::property_set::Guid,
        property_identifier: u32,
    ) -> Result<&mut Self, OleError> {
        self.set_property(
            format_identifier,
            property_identifier,
            Value::Storage(IndirectPropertyName::new(property_identifier)?),
        )
    }

    pub fn set_streamed_object_property(
        &mut self,
        format_identifier: crate::property_set::Guid,
        property_identifier: u32,
    ) -> Result<&mut Self, OleError> {
        self.set_property(
            format_identifier,
            property_identifier,
            Value::StreamedObject(IndirectPropertyName::new(property_identifier)?),
        )
    }

    pub fn set_stored_object_property(
        &mut self,
        format_identifier: crate::property_set::Guid,
        property_identifier: u32,
    ) -> Result<&mut Self, OleError> {
        self.set_property(
            format_identifier,
            property_identifier,
            Value::StoredObject(IndirectPropertyName::new(property_identifier)?),
        )
    }

    pub fn set_versioned_stream_property(
        &mut self,
        format_identifier: crate::property_set::Guid,
        property_identifier: u32,
        version_guid: crate::property_set::Guid,
    ) -> Result<&mut Self, OleError> {
        self.set_property(
            format_identifier,
            property_identifier,
            Value::VersionedStream(VersionedStream::new(version_guid, property_identifier)?),
        )
    }

    pub fn add_storage(
        &mut self,
        path: &[&str],
        class_id: Option<crate::property_set::Guid>,
    ) -> Result<&mut Self, OleError> {
        let records = self.records();
        validate_new_path(&self.source.state, records, path)?;
        let mut candidate_records =
            clone_records(records, self.source.state.limits.max_total_path_bytes)?;
        let owned_path = canonical_new_path(records, path, "non-simple storage path")?;
        let metadata = Metadata::staged_storage(class_id);
        candidate_records
            .try_reserve(1)
            .map_err(|source| OleError::Allocation {
                resource: "non-simple storage records",
                source,
            })?;
        candidate_records.push(EntryRecord {
            path: owned_path,
            kind: Some(ElementKind::Storage),
            metadata: Some(metadata),
            raw_kind: EntryKind::Storage.raw(),
            size: 0,
        });
        validate_records(&self.source.state, &candidate_records)?;
        self.records = Some(candidate_records);
        Ok(self)
    }

    pub fn set_stream(&mut self, path: &[&str], bytes: Vec<u8>) -> Result<&mut Self, OleError> {
        if path.len() == 1 && directory_names_equal(path[0], "CONTENTS") {
            return Err(invalid("use replace_contents to edit the CONTENTS stream"));
        }
        if bytes.len() as u64 > self.source.state.limits.max_stream_bytes {
            return Err(OleError::LimitExceeded {
                resource: "non-simple Property Set stream bytes",
                observed: bytes.len() as u64,
                maximum: self.source.state.limits.max_stream_bytes,
            });
        }
        let records = self.records();
        let existing = records
            .iter()
            .find(|record| path_matches(path, &record.path));
        if let Some(record) = existing
            && record.kind != Some(ElementKind::Stream)
        {
            return Err(invalid("non-simple path identifies a storage"));
        }
        let owned_path = if existing.is_none() {
            validate_new_path(&self.source.state, records, path)?;
            Some(canonical_new_path(records, path, "non-simple stream path")?)
        } else {
            None
        };
        let mut candidate_records =
            clone_records(records, self.source.state.limits.max_total_path_bytes)?;
        let mut candidate_edits =
            clone_edits(&self.edits, self.source.state.limits.max_total_path_bytes)?;
        if let Some(record) = candidate_records
            .iter_mut()
            .find(|record| path_matches(path, &record.path))
        {
            record.size = bytes.len() as u64;
            stage_edit(
                &mut candidate_edits,
                &record.path,
                Some(bytes.into()),
                self.source.state.limits.max_total_path_bytes,
            )?;
        } else {
            let owned_path = owned_path.ok_or_else(|| {
                invalid("non-simple stream path disappeared during transaction planning")
            })?;
            candidate_records
                .try_reserve(1)
                .map_err(|source| OleError::Allocation {
                    resource: "non-simple stream records",
                    source,
                })?;
            let size = bytes.len() as u64;
            candidate_records.push(EntryRecord {
                path: owned_path.clone(),
                kind: Some(ElementKind::Stream),
                metadata: None,
                raw_kind: EntryKind::Stream.raw(),
                size,
            });
            stage_edit(
                &mut candidate_edits,
                &owned_path,
                Some(bytes.into()),
                self.source.state.limits.max_total_path_bytes,
            )?;
        }
        validate_records(&self.source.state, &candidate_records)?;
        self.records = Some(candidate_records);
        self.edits = candidate_edits;
        Ok(self)
    }

    pub fn remove_element(&mut self, path: &[&str]) -> Result<&mut Self, OleError> {
        if path.is_empty() || (path.len() == 1 && directory_names_equal(path[0], "CONTENTS")) {
            return Err(invalid(
                "the non-simple root and CONTENTS stream are required",
            ));
        }
        let records = self.records();
        let Some(target) = records
            .iter()
            .find(|record| path_matches(path, &record.path))
        else {
            return Err(OleError::StreamNotFound);
        };
        let target_path = clone_path(&target.path, "non-simple removal path")?;
        let mut candidate_records =
            clone_records(records, self.source.state.limits.max_total_path_bytes)?;
        candidate_records.retain(|record| !has_path_prefix(&record.path, &target_path));
        let mut candidate_edits =
            clone_edits(&self.edits, self.source.state.limits.max_total_path_bytes)?;
        candidate_edits.retain(|edit| !has_path_prefix(&edit.path, &target_path));
        validate_records(&self.source.state, &candidate_records)?;
        self.records = Some(candidate_records);
        self.edits = candidate_edits;
        Ok(self)
    }

    pub fn snapshot(&self) -> Result<Snapshot, OleError> {
        self.materialize()
    }

    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    pub fn commit(self) -> Result<Commit, OleError> {
        let snapshot = self.materialize()?;
        let patch = Patch::new(&self.source, &snapshot);
        Ok(Commit { snapshot, patch })
    }

    fn ensure_contents(&self) -> Result<Stream, OleError> {
        match &self.candidate {
            Some(candidate) => clone_stream(candidate),
            None => self.source.contents(),
        }
    }

    fn set_candidate(&mut self, candidate: Stream) -> Result<(), OleError> {
        let records = self
            .records
            .as_deref()
            .unwrap_or(&self.source.state.records);
        validate_records(&self.source.state, records)?;
        codec::validate_closure(&self.source.state, &candidate, records)?;
        self.candidate = Some(candidate);
        Ok(())
    }

    fn records(&self) -> &[EntryRecord] {
        self.records
            .as_deref()
            .unwrap_or(&self.source.state.records)
    }

    fn materialize(&self) -> Result<Snapshot, OleError> {
        if self.candidate.is_none() && self.records.is_none() && self.edits.is_empty() {
            return Ok(self.source.clone());
        }
        let records = self
            .records
            .as_deref()
            .unwrap_or(&self.source.state.records);
        validate_records(&self.source.state, records)?;
        match self.candidate.as_ref() {
            // Validate and serialize the borrowed candidate before cloning its
            // section/value tree.  A caller-controlled CONTENTS limit must
            // reject an oversized candidate while it is still source-backed.
            Some(candidate) => self.materialize_candidate(candidate, records),
            None => {
                let candidate = self.source.contents()?;
                self.materialize_candidate(&candidate, records)
            },
        }
    }

    fn materialize_candidate(
        &self,
        candidate: &Stream,
        records: &[EntryRecord],
    ) -> Result<Snapshot, OleError> {
        codec::validate_closure(&self.source.state, candidate, records)?;
        let contents_bytes =
            match candidate.to_bytes_with_limit(self.source.state.limits.max_contents_bytes) {
                Err(OleError::LimitExceeded {
                    observed, maximum, ..
                }) => {
                    return Err(OleError::LimitExceeded {
                        resource: "non-simple Property Set CONTENTS bytes",
                        observed,
                        maximum,
                    });
                },
                result => result?,
            };
        if self.is_exact_source(records, &contents_bytes)? {
            return Ok(self.source.clone());
        }

        // All candidate-owned admission work is complete.  Only changed
        // publication needs a retained owned tree for the renderer.
        let owned_candidate = clone_stream(candidate)?;
        let bytes = codec::render(
            &self.source.state,
            records,
            &owned_candidate,
            contents_bytes.as_slice(),
            &self.edits,
        )?;
        Snapshot::open_shared(bytes, self.source.state.limits)
    }

    fn is_exact_source(
        &self,
        records: &[EntryRecord],
        contents_bytes: &[u8],
    ) -> Result<bool, OleError> {
        if records.len() != self.source.state.records.len()
            || records
                .iter()
                .zip(self.source.state.records.iter())
                .any(|(left, right)| {
                    left.path != right.path
                        || left.kind != right.kind
                        || left.metadata != right.metadata
                        || left.raw_kind != right.raw_kind
                        || left.size != right.size
                })
        {
            return Ok(false);
        }
        let source_contents = codec::read_contents_bytes(&self.source.state)?;
        if source_contents.as_ref() != contents_bytes {
            return Ok(false);
        }
        for edit in &self.edits {
            let refs = path_refs(&edit.path)?;
            let Some(record) = codec::find_record(&self.source.state.records, &refs) else {
                return Ok(false);
            };
            let Some(bytes) = edit.bytes.as_ref() else {
                return Ok(false);
            };
            if codec::read_stream(&self.source.state, record)?.as_ref() != bytes.as_ref() {
                return Ok(false);
            }
        }
        Ok(true)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    limits: Limits,
}

impl Patch {
    fn new(before: &Snapshot, after: &Snapshot) -> Self {
        Self {
            base: Revision::of(&before.state.source),
            target: Revision::of(&after.state.source),
            before: before.source_shared(),
            after: after.source_shared(),
            limits: after.limits(),
        }
    }

    #[must_use]
    pub const fn base(&self) -> Revision {
        self.base
    }

    #[must_use]
    pub const fn target(&self) -> Revision {
        self.target
    }

    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        &self.before
    }

    #[must_use]
    pub fn before(&self) -> &[u8] {
        self.before_bytes()
    }

    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        &self.after
    }

    #[must_use]
    pub fn after(&self) -> &[u8] {
        self.after_bytes()
    }

    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.base.value()
    }

    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target.value()
    }

    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_ref() == self.after.as_ref()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.is_noop()
    }

    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, OleError> {
        if Revision::of(&source.state.source) != self.base
            || source.state.source.as_ref() != self.before.as_ref()
        {
            return Err(invalid(
                "non-simple Property Set patch source does not match its base",
            ));
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        Snapshot::open_shared(Arc::clone(&self.after), self.limits)
    }

    pub fn revert(&self, target: &Snapshot) -> Result<Snapshot, OleError> {
        self.inverse().apply(target)
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            limits: self.limits,
        }
    }
}

pub fn update<F>(snapshot: &Snapshot, edit: F) -> Result<Commit, OleError>
where
    F: FnOnce(&mut Editor) -> Result<(), OleError>,
{
    let mut editor = snapshot.edit();
    edit(&mut editor)?;
    editor.commit()
}

fn clone_stream(source: &Stream) -> Result<Stream, OleError> {
    let mut sections = Vec::new();
    sections
        .try_reserve_exact(source.sections.len())
        .map_err(|source| OleError::Allocation {
            resource: "non-simple Property Set sections",
            source,
        })?;
    for section in &source.sections {
        sections.push(try_clone_property_set(section)?);
    }
    Ok(Stream {
        version: source.version,
        system_identifier: source.system_identifier,
        class_identifier: source.class_identifier,
        sections,
    })
}

fn clone_records(
    source: &[EntryRecord],
    max_path_bytes: usize,
) -> Result<Vec<EntryRecord>, OleError> {
    let mut path_bytes = 0usize;
    for record in source {
        for component in &record.path {
            path_bytes = path_bytes
                .checked_add(component.len())
                .ok_or_else(|| invalid("non-simple Property Set path bytes overflow"))?;
        }
    }
    if path_bytes > max_path_bytes {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set copied path bytes",
            observed: path_bytes as u64,
            maximum: max_path_bytes as u64,
        });
    }
    let mut records = Vec::new();
    records
        .try_reserve_exact(source.len())
        .map_err(|source| OleError::Allocation {
            resource: "non-simple transaction directory records",
            source,
        })?;
    for record in source {
        let mut path = Vec::new();
        path.try_reserve_exact(record.path.len())
            .map_err(|source| OleError::Allocation {
                resource: "non-simple transaction directory path",
                source,
            })?;
        path.extend(record.path.iter().cloned());
        records.push(EntryRecord {
            path,
            kind: record.kind,
            metadata: record.metadata,
            raw_kind: record.raw_kind,
            size: record.size,
        });
    }
    Ok(records)
}

fn clone_edits(
    source: &[PayloadEdit],
    max_path_bytes: usize,
) -> Result<Vec<PayloadEdit>, OleError> {
    let path_bytes = path_bytes(source.iter().map(|edit| edit.path.as_slice()))?;
    if path_bytes > max_path_bytes {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set copied path bytes",
            observed: path_bytes as u64,
            maximum: max_path_bytes as u64,
        });
    }
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(source.len())
        .map_err(|source| OleError::Allocation {
            resource: "non-simple payload edits",
            source,
        })?;
    for edit in source {
        let mut path = Vec::new();
        path.try_reserve_exact(edit.path.len())
            .map_err(|source| OleError::Allocation {
                resource: "non-simple payload edit path",
                source,
            })?;
        path.extend(edit.path.iter().cloned());
        edits.push(PayloadEdit {
            path,
            bytes: edit.bytes.as_ref().map(Arc::clone),
        });
    }
    Ok(edits)
}

fn clone_path<T: AsRef<str>>(path: &[T], resource: &'static str) -> Result<Vec<String>, OleError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(path.len())
        .map_err(|source| OleError::Allocation { resource, source })?;
    for component in path {
        let component = component.as_ref();
        validate_directory_name(component)?;
        output
            .try_reserve(component.len())
            .map_err(|source| OleError::Allocation { resource, source })?;
        output.push(component.to_owned());
    }
    Ok(output)
}

fn path_bytes<'a>(mut paths: impl Iterator<Item = &'a [String]>) -> Result<usize, OleError> {
    paths.try_fold(0usize, |total, path| {
        path.iter().try_fold(total, |total, component| {
            total
                .checked_add(component.len())
                .ok_or_else(|| invalid("non-simple Property Set path bytes overflow"))
        })
    })
}

fn canonical_new_path(
    records: &[EntryRecord],
    path: &[&str],
    resource: &'static str,
) -> Result<Vec<String>, OleError> {
    let mut owned = clone_path(path, resource)?;
    for depth in 0..path.len() {
        let Some(existing) = records.iter().find(|record| {
            record.path.len() == depth + 1
                && record.path[..depth]
                    .iter()
                    .zip(&path[..depth])
                    .all(|(left, right)| directory_names_equal(left, right))
                && directory_names_equal(&record.path[depth], path[depth])
        }) else {
            continue;
        };
        owned[depth].clear();
        owned[depth]
            .try_reserve_exact(existing.path[depth].len())
            .map_err(|source| OleError::Allocation { resource, source })?;
        owned[depth].push_str(&existing.path[depth]);
    }
    Ok(owned)
}

fn path_matches(path: &[&str], candidate: &[String]) -> bool {
    path.len() == candidate.len()
        && path
            .iter()
            .zip(candidate)
            .all(|(left, right)| directory_names_equal(left, right))
}

fn has_path_prefix(path: &[String], prefix: &[String]) -> bool {
    path.len() >= prefix.len()
        && path
            .iter()
            .zip(prefix)
            .all(|(left, right)| directory_names_equal(left, right))
}

fn owned_path_matches(left: &[String], right: &[String]) -> bool {
    left.len() == right.len()
        && left
            .iter()
            .zip(right)
            .all(|(left, right)| directory_names_equal(left, right))
}

fn validate_new_path(
    state: &State,
    records: &[EntryRecord],
    path: &[&str],
) -> Result<(), OleError> {
    if path.is_empty() {
        return Err(invalid("non-simple CFB path cannot be empty"));
    }
    for component in path {
        validate_directory_name(component)?;
    }
    if path
        .iter()
        .any(|component| component.len() > state.limits.max_name_bytes)
    {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set directory name bytes",
            observed: path
                .iter()
                .map(|component| component.len())
                .max()
                .unwrap_or(0) as u64,
            maximum: state.limits.max_name_bytes as u64,
        });
    }
    if path.len() > state.limits.max_storage_depth.saturating_add(1) {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set storage depth",
            observed: path.len() as u64,
            maximum: state.limits.max_storage_depth.saturating_add(1) as u64,
        });
    }
    if records
        .iter()
        .any(|record| path_matches(path, &record.path))
    {
        return Err(invalid("non-simple CFB path already exists"));
    }
    if path.len() > 1 {
        let parent = &path[..path.len() - 1];
        let parent_record = records
            .iter()
            .find(|record| path_matches(parent, &record.path));
        if parent_record.is_none_or(|record| record.kind != Some(ElementKind::Storage)) {
            return Err(invalid("non-simple CFB path parent is not a storage"));
        }
    }
    Ok(())
}

fn stage_edit(
    edits: &mut Vec<PayloadEdit>,
    path: &[String],
    bytes: Option<Arc<[u8]>>,
    max_path_bytes: usize,
) -> Result<(), OleError> {
    if let Some(edit) = edits.iter_mut().find(|edit| same_path(&edit.path, path)) {
        edit.bytes = bytes;
        return Ok(());
    }
    let current_path_bytes = path_bytes(edits.iter().map(|edit| edit.path.as_slice()))?;
    let added_path_bytes = path_bytes(std::iter::once(path))?;
    let total_path_bytes = current_path_bytes
        .checked_add(added_path_bytes)
        .ok_or_else(|| invalid("non-simple Property Set path bytes overflow"))?;
    if total_path_bytes > max_path_bytes {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set copied path bytes",
            observed: total_path_bytes as u64,
            maximum: max_path_bytes as u64,
        });
    }
    let mut owned_path = Vec::new();
    owned_path
        .try_reserve_exact(path.len())
        .map_err(|source| OleError::Allocation {
            resource: "non-simple payload edit path",
            source,
        })?;
    owned_path.extend(path.iter().cloned());
    edits
        .try_reserve(1)
        .map_err(|source| OleError::Allocation {
            resource: "non-simple payload edits",
            source,
        })?;
    edits.push(PayloadEdit {
        path: owned_path,
        bytes,
    });
    Ok(())
}

fn validate_records(state: &State, records: &[EntryRecord]) -> Result<(), OleError> {
    if records.is_empty() || records.len() > state.limits.max_directory_entries {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set directory entries",
            observed: records.len() as u64,
            maximum: state.limits.max_directory_entries as u64,
        });
    }
    if !records.iter().any(EntryRecord::is_root) {
        return Err(invalid("non-simple CFB candidate has no root entry"));
    }
    let mut total_names = 0usize;
    let mut total_path_bytes = 0usize;
    let mut streams = 0usize;
    let mut total_stream_bytes = 0u64;
    for record in records {
        if record.path.len() > state.limits.max_storage_depth.saturating_add(1) {
            return Err(OleError::LimitExceeded {
                resource: "non-simple Property Set storage depth",
                observed: record.path.len() as u64,
                maximum: state.limits.max_storage_depth.saturating_add(1) as u64,
            });
        }
        for component in &record.path {
            validate_directory_name(component)?;
            if component.len() > state.limits.max_name_bytes {
                return Err(OleError::LimitExceeded {
                    resource: "non-simple Property Set directory name bytes",
                    observed: component.len() as u64,
                    maximum: state.limits.max_name_bytes as u64,
                });
            }
            total_names = total_names
                .checked_add(component.len())
                .ok_or_else(|| invalid("non-simple Property Set path bytes overflow"))?;
            total_path_bytes = total_path_bytes
                .checked_add(component.len())
                .ok_or_else(|| invalid("non-simple Property Set path bytes overflow"))?;
        }
        if record.kind == Some(ElementKind::Stream) {
            streams = streams
                .checked_add(1)
                .ok_or_else(|| invalid("non-simple Property Set stream count overflow"))?;
            if record.size > state.limits.max_stream_bytes {
                return Err(OleError::LimitExceeded {
                    resource: "non-simple Property Set stream bytes",
                    observed: record.size,
                    maximum: state.limits.max_stream_bytes,
                });
            }
            total_stream_bytes = total_stream_bytes
                .checked_add(record.size)
                .ok_or_else(|| invalid("non-simple Property Set stream bytes overflow"))?;
        }
    }
    if total_names > state.limits.max_total_name_bytes {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set directory metadata bytes",
            observed: total_names as u64,
            maximum: state.limits.max_total_name_bytes as u64,
        });
    }
    if total_path_bytes > state.limits.max_total_path_bytes {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set copied path bytes",
            observed: total_path_bytes as u64,
            maximum: state.limits.max_total_path_bytes as u64,
        });
    }
    if streams > state.limits.max_streams {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set streams",
            observed: streams as u64,
            maximum: state.limits.max_streams as u64,
        });
    }
    if total_stream_bytes > state.limits.max_total_stream_bytes {
        return Err(OleError::LimitExceeded {
            resource: "non-simple Property Set aggregate stream bytes",
            observed: total_stream_bytes,
            maximum: state.limits.max_total_stream_bytes,
        });
    }
    for record in records.iter().filter(|record| !record.is_root()) {
        let parent = &record.path[..record.path.len() - 1];
        let Some(parent_record) = records
            .iter()
            .find(|candidate| owned_path_matches(&candidate.path, parent))
        else {
            return Err(invalid("non-simple CFB record has a missing parent"));
        };
        if parent.is_empty() {
            if parent_record.kind.is_some() {
                return Err(invalid("non-simple CFB root record has the wrong kind"));
            }
        } else if parent_record.kind != Some(ElementKind::Storage) {
            return Err(invalid("non-simple CFB record parent is not a storage"));
        }
    }
    Ok(())
}

fn path_refs(path: &[String]) -> Result<Vec<&str>, OleError> {
    let mut refs = Vec::new();
    refs.try_reserve_exact(path.len())
        .map_err(|source| OleError::Allocation {
            resource: "non-simple patch path references",
            source,
        })?;
    refs.extend(path.iter().map(String::as_str));
    Ok(refs)
}
