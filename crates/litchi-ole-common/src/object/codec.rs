//! CFB capture and deterministic rendering for the object owner.

use super::cfb_path::{CfbPath, path_identity_hash, same_path};
use super::directory::{self, EntryKind};
use super::model::{Limits, Object, Storage, Stream};
use super::target::Target;
use crate::property_set::Guid;
use crate::protection::reject_protected_container;
use litchi_cfb::{
    OleError, OleFile, OleWriter, OverlayError, OverlayLimits, SameLengthStreamOverlay,
    SectorLayoutPolicy, SharedOleFile,
};
use litchi_core::SourceVersion;
use std::collections::HashMap;
use std::io::{Cursor, Read, Seek};
use std::sync::Arc;

#[derive(Debug, Clone)]
pub(crate) struct Package {
    sector_size: usize,
    root_clsid: Option<Guid>,
    root_directory: Option<directory::Metadata>,
    storages: Vec<Storage>,
    streams: Vec<Stream>,
}

impl Package {
    pub(crate) fn capture<R: Read + Seek>(
        ole: &mut OleFile<R>,
        limits: Limits,
    ) -> Result<Self, OleError> {
        Self::capture_with_budget(
            ole,
            limits,
            limits.max_package_storages(),
            limits.max_streams,
            limits.max_total_size,
            true,
        )
    }

    /// Captures a standalone replacement/object compound file.  The file is
    /// one selected object, so its storage ceiling is the per-object depth
    /// limit rather than the aggregate package product.
    pub(crate) fn capture_object<R: Read + Seek>(
        ole: &mut OleFile<R>,
        limits: Limits,
    ) -> Result<Self, OleError> {
        Self::capture_with_budget(
            ole,
            limits,
            limits.max_object_storages(),
            limits.max_object_streams(),
            limits.max_object_size,
            false,
        )
    }

    fn capture_with_budget<R: Read + Seek>(
        ole: &mut OleFile<R>,
        limits: Limits,
        max_storages: usize,
        max_streams: usize,
        max_bytes: u64,
        retain_root: bool,
    ) -> Result<Self, OleError> {
        let root_directory = ole.root_entry().map(directory::decode).transpose()?;
        let mut package = Self {
            sector_size: ole.sector_size(),
            root_clsid: root_directory.and_then(directory::Metadata::class_id),
            root_directory,
            storages: Vec::new(),
            streams: Vec::new(),
        };
        if !retain_root {
            package.root_directory = None;
        }
        let mut budget = Budget::new(max_storages, max_streams, max_bytes);
        capture_container(ole, &[], &mut package, &mut budget, limits)?;
        Ok(package)
    }

    pub(crate) fn capture_target<R: Read + Seek>(
        ole: &mut OleFile<R>,
        target: &Target,
        limits: Limits,
    ) -> Result<Object, OleError> {
        if target.path().len() > limits.max_storage_depth {
            return Err(OleError::InvalidFormat(
                "object target path exceeds storage depth limit".into(),
            ));
        }
        let resolved_target = target.resolve(ole, limits.max_package_directory_entries())?;
        let storage = find_storage(
            ole,
            resolved_target.path(),
            limits.max_package_directory_entries(),
        )?;
        let mut package = Self {
            sector_size: ole.sector_size(),
            root_clsid: storage.class_id(),
            // A selected storage is promoted into a standalone compound
            // file.  Its directory metadata remains attached to the
            // selected Object, but it is not the standalone file's root
            // directory metadata.
            root_directory: None,
            storages: Vec::new(),
            streams: Vec::new(),
        };
        let mut budget = Budget::new(
            limits.max_object_storages(),
            limits.max_object_streams(),
            limits.max_object_size,
        );
        capture_subtree(
            ole,
            resolved_target.path(),
            &[],
            &mut package,
            &mut budget,
            limits,
        )?;
        package.object_from_root(resolved_target, storage, limits)
    }

    pub(crate) fn object(&self, target: Target, limits: Limits) -> Result<Object, OleError> {
        let storage = self
            .storages
            .iter()
            .find(|storage| storage.path() == target.path())
            .cloned()
            .ok_or_else(|| {
                OleError::InvalidFormat(format!("object storage {:?} not found", target.path()))
            })?;
        let object_package = Self {
            sector_size: self.sector_size,
            root_clsid: storage.class_id(),
            // The selected storage is promoted to a fresh standalone root.
            // Keep the source storage metadata on `Object::storage()` rather
            // than projecting storage-only fields onto the new root.
            root_directory: None,
            storages: self
                .storages
                .iter()
                .filter(|value| {
                    value.path().len() > target.path().len()
                        && value.path().starts_with(target.path())
                })
                .map(|value| {
                    Storage::new(
                        value.path()[target.path().len()..].to_vec(),
                        *value.directory(),
                    )
                })
                .collect(),
            streams: self
                .streams
                .iter()
                .filter(|value| {
                    value.path().len() > target.path().len()
                        && value.path().starts_with(target.path())
                })
                .map(|value| {
                    Stream::new(
                        value.path()[target.path().len()..].to_vec(),
                        value.bytes_shared(),
                        value.directory().copied(),
                    )
                })
                .collect(),
        };
        object_package.object_from_root(target, storage, limits)
    }

    pub(crate) fn put_stream(
        &mut self,
        path: &[String],
        data: Arc<[u8]>,
        limits: Limits,
    ) -> Result<(), OleError> {
        if data.len() as u64 > limits.max_stream_size {
            return Err(OleError::InvalidFormat(
                "replacement stream exceeds size limit".into(),
            ));
        }
        let stream = self
            .streams
            .iter_mut()
            .find(|stream| stream.path() == path)
            .ok_or(OleError::StreamNotFound)?;
        // Replacing bytes does not edit the stream's directory metadata. Keep
        // that projection available so an equal-length replacement can use
        // the validated physical overlay; length changes still decline that
        // path and let the source-layout writer patch start/size fields.
        let directory = stream.directory().copied();
        *stream = Stream::new(path.to_vec(), data, directory);
        self.check(limits)
    }

    pub(crate) fn stream(&self, path: &[String]) -> Option<&[u8]> {
        self.streams
            .iter()
            .find(|stream| stream.path() == path)
            .map(Stream::bytes)
    }

    pub(crate) fn stream_shared(&self, path: &[String]) -> Option<Arc<[u8]>> {
        self.streams
            .iter()
            .find(|stream| stream.path() == path)
            .map(Stream::bytes_shared)
    }

    pub(crate) fn reuse_stream_allocations(&mut self, previous: &Self) -> Result<(), OleError> {
        let mut by_path = HashMap::new();
        by_path
            .try_reserve(previous.streams.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB stream allocation index",
                source,
            })?;
        for stream in &previous.streams {
            by_path.insert(stream.path(), stream);
        }
        for stream in &mut self.streams {
            if let Some(cached_stream) = by_path.get(stream.path())
                && cached_stream.bytes() == stream.bytes()
            {
                stream.replace_data(cached_stream.bytes_shared());
            }
        }
        Ok(())
    }

    pub(crate) fn add_stream(
        &mut self,
        path: Vec<String>,
        data: Arc<[u8]>,
        limits: Limits,
    ) -> Result<(), OleError> {
        if path.is_empty() || path.iter().any(String::is_empty) {
            return Err(OleError::InvalidFormat(
                "new package stream path must contain names".into(),
            ));
        }
        if data.len() as u64 > limits.max_stream_size {
            return Err(OleError::InvalidFormat(
                "new package stream exceeds size limit".into(),
            ));
        }
        if self.streams.iter().any(|stream| stream.path() == path) {
            return Err(OleError::InvalidFormat(format!(
                "package stream {path:?} already exists"
            )));
        }
        if path.len() > 1
            && !self
                .storages
                .iter()
                .any(|storage| storage.path() == &path[..path.len() - 1])
        {
            return Err(OleError::InvalidFormat(
                "new package stream parent storage is missing".into(),
            ));
        }
        self.streams.push(Stream::new(path, data, None));
        self.check(limits)
    }

    pub(crate) fn remove_stream(
        &mut self,
        path: &CfbPath,
        limits: Limits,
    ) -> Result<Option<Arc<[u8]>>, OleError> {
        let Some(removed) = self.removable_stream(path)? else {
            return Ok(None);
        };
        let stream = self
            .streams
            .iter()
            .position(|stream| same_path(stream.path(), path.as_slice()));
        let stream = stream.ok_or_else(|| {
            OleError::InvalidFormat("resolved CFB stream disappeared before removal".into())
        })?;
        self.streams.remove(stream);
        self.check(limits)?;
        Ok(Some(removed))
    }

    pub(crate) fn removable_stream(&self, path: &CfbPath) -> Result<Option<Arc<[u8]>>, OleError> {
        if let Some(stream) = self
            .streams
            .iter()
            .find(|stream| same_path(stream.path(), path.as_slice()))
        {
            return Ok(Some(stream.bytes_shared()));
        }
        if self
            .storages
            .iter()
            .any(|storage| same_path(storage.path(), path.as_slice()))
        {
            return Err(OleError::InvalidFormat(format!(
                "package entry {:?} is a storage, not a stream",
                path.as_slice()
            )));
        }
        Ok(None)
    }

    pub(crate) fn remove_streams<'a>(
        &mut self,
        paths: impl ExactSizeIterator<Item = &'a CfbPath>,
        limits: Limits,
    ) -> Result<Vec<Option<Arc<[u8]>>>, OleError> {
        let mut by_identity = HashMap::new();
        by_identity
            .try_reserve(self.streams.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB stream removal index",
                source,
            })?;
        for (index, stream) in self.streams.iter().enumerate() {
            by_identity.insert(path_identity_hash(stream.path()), index);
        }
        let mut storage_by_identity = HashMap::new();
        storage_by_identity
            .try_reserve(self.storages.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB storage identity index",
                source,
            })?;
        for (index, storage) in self.storages.iter().enumerate() {
            storage_by_identity.insert(path_identity_hash(storage.path()), index);
        }

        let mut removed = Vec::new();
        removed
            .try_reserve_exact(paths.len())
            .map_err(|source| OleError::Allocation {
                resource: "stream removal results",
                source,
            })?;
        let mut positions = Vec::new();
        positions
            .try_reserve_exact(paths.len())
            .map_err(|source| OleError::Allocation {
                resource: "stream removal positions",
                source,
            })?;
        for path in paths {
            let identity = path.identity_hash();
            let position = match by_identity.get(&identity).copied() {
                Some(index) if same_path(self.streams[index].path(), path.as_slice()) => {
                    Some(index)
                },
                Some(_collision) => self
                    .streams
                    .iter()
                    .position(|stream| same_path(stream.path(), path.as_slice())),
                None => None,
            };
            if let Some(position) = position {
                removed.push(Some(self.streams[position].bytes_shared()));
                positions.push(position);
            } else {
                let is_storage = match storage_by_identity.get(&identity).copied() {
                    Some(index) if same_path(self.storages[index].path(), path.as_slice()) => true,
                    Some(_collision) => self
                        .storages
                        .iter()
                        .any(|storage| same_path(storage.path(), path.as_slice())),
                    None => false,
                };
                if is_storage {
                    return Err(OleError::InvalidFormat(format!(
                        "package entry {:?} is a storage, not a stream",
                        path.as_slice()
                    )));
                }
                removed.push(None);
            }
        }

        if !positions.is_empty() {
            let mut selected = Vec::new();
            selected
                .try_reserve_exact(self.streams.len())
                .map_err(|source| OleError::Allocation {
                    resource: "CFB stream removal bitmap",
                    source,
                })?;
            selected.resize(self.streams.len(), false);
            for position in positions {
                selected[position] = true;
            }
            let mut position = 0usize;
            self.streams.retain(|_| {
                let keep = !selected[position];
                position += 1;
                keep
            });
            self.check(limits)?;
        }
        Ok(removed)
    }

    pub(crate) fn replace_object(
        &mut self,
        path: &[String],
        replacement: &Self,
        limits: Limits,
    ) -> Result<(), OleError> {
        limits.validate()?;
        let root_index = self
            .storages
            .iter()
            .position(|storage| storage.path() == path)
            .ok_or_else(|| OleError::InvalidFormat(format!("object storage {path:?} not found")))?;
        let removed_storages = self
            .storages
            .iter()
            .filter(|storage| storage.path() != path && storage.path().starts_with(path))
            .count();
        let storage_count = self
            .storages
            .len()
            .checked_sub(removed_storages)
            .and_then(|count| count.checked_add(replacement.storages.len()))
            .ok_or_else(|| OleError::InvalidFormat("CFB storage count overflow".into()))?;
        let removed_bytes = stream_bytes_total(
            self.streams
                .iter()
                .filter(|stream| stream.path().starts_with(path)),
        )?;
        let current_bytes = stream_bytes_total(self.streams.iter())?;
        let replacement_bytes = stream_bytes_total(replacement.streams.iter())?;
        let total_bytes = current_bytes
            .checked_sub(removed_bytes)
            .and_then(|value| value.checked_add(replacement_bytes))
            .ok_or_else(|| OleError::InvalidFormat("CFB capture size overflow".into()))?;
        let stream_count = self
            .streams
            .len()
            .checked_sub(
                self.streams
                    .iter()
                    .filter(|stream| stream.path().starts_with(path))
                    .count(),
            )
            .and_then(|count| count.checked_add(replacement.streams.len()))
            .ok_or_else(|| OleError::InvalidFormat("CFB stream count overflow".into()))?;
        check_counts(limits, storage_count, stream_count, total_bytes)?;

        // Build every replacement path before changing the current package.
        // This keeps path and vector allocation failures outside the mutation
        // window, while the count/byte preflight above keeps limit failures
        // atomic as well.
        let root_directory = self.storages[root_index]
            .directory()
            .with_class_id(replacement.root_clsid);
        let replacement_root = Storage::new(clone_path(path)?, root_directory);
        let mut replacement_storages = Vec::new();
        replacement_storages
            .try_reserve_exact(replacement.storages.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB replacement storages",
                source,
            })?;
        for storage in &replacement.storages {
            replacement_storages.push(Storage::new(
                join_checked(path, storage.path())?,
                *storage.directory(),
            ));
        }
        let mut replacement_streams = Vec::new();
        replacement_streams
            .try_reserve_exact(replacement.streams.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB replacement streams",
                source,
            })?;
        for stream in &replacement.streams {
            replacement_streams.push(Stream::new(
                join_checked(path, stream.path())?,
                stream.bytes_shared(),
                stream.directory().copied(),
            ));
        }
        self.storages
            .try_reserve_exact(replacement_storages.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB replacement storage capacity",
                source,
            })?;
        self.streams
            .try_reserve_exact(replacement_streams.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB replacement stream capacity",
                source,
            })?;

        // Replacing an object changes its class identity intentionally, while
        // the existing target storage remains the same directory object in
        // the containing file.  Preserve its state and timestamps by default;
        // replacement root metadata belongs to the standalone source file.
        self.storages[root_index] = replacement_root;
        self.storages.retain(|storage| {
            storage.path() == path
                || !(storage.path().len() > path.len() && storage.path().starts_with(path))
        });
        self.streams
            .retain(|stream| !stream.path().starts_with(path));
        self.storages.extend(replacement_storages);
        self.streams.extend(replacement_streams);
        Ok(())
    }

    pub(crate) fn add_object(
        &mut self,
        target: &Target,
        replacement: &Self,
        limits: Limits,
    ) -> Result<(), OleError> {
        limits.validate()?;
        if self
            .storages
            .iter()
            .any(|storage| storage.path() == target.path())
        {
            return Err(OleError::InvalidFormat(format!(
                "object storage {:?} already exists",
                target.path()
            )));
        }
        if target.path().len() > 1
            && !self
                .storages
                .iter()
                .any(|storage| storage.path() == &target.path()[..target.path().len() - 1])
        {
            return Err(OleError::InvalidFormat(
                "new object storage parent is missing".into(),
            ));
        }
        let storage_count = self
            .storages
            .len()
            .checked_add(1)
            .and_then(|count| count.checked_add(replacement.storages.len()))
            .ok_or_else(|| OleError::InvalidFormat("CFB storage count overflow".into()))?;
        let stream_count = self
            .streams
            .len()
            .checked_add(replacement.streams.len())
            .ok_or_else(|| OleError::InvalidFormat("CFB stream count overflow".into()))?;
        let total_bytes = stream_bytes_total(self.streams.iter())?
            .checked_add(stream_bytes_total(replacement.streams.iter())?)
            .ok_or_else(|| OleError::InvalidFormat("CFB capture size overflow".into()))?;
        check_counts(limits, storage_count, stream_count, total_bytes)?;

        let root_directory = directory::Metadata::staged_storage(replacement.root_clsid);
        // A newly added target is a fresh storage in the containing file.
        // Use normative storage defaults; source root state/times are not
        // meaningful storage metadata.
        let mut added_storages = Vec::new();
        added_storages
            .try_reserve_exact(replacement.storages.len().saturating_add(1))
            .map_err(|source| OleError::Allocation {
                resource: "CFB added storages",
                source,
            })?;
        added_storages.push(Storage::new(clone_path(target.path())?, root_directory));
        for storage in &replacement.storages {
            added_storages.push(Storage::new(
                join_checked(target.path(), storage.path())?,
                *storage.directory(),
            ));
        }
        let mut added_streams = Vec::new();
        added_streams
            .try_reserve_exact(replacement.streams.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB added streams",
                source,
            })?;
        for stream in &replacement.streams {
            added_streams.push(Stream::new(
                join_checked(target.path(), stream.path())?,
                stream.bytes_shared(),
                stream.directory().copied(),
            ));
        }
        self.storages
            .try_reserve_exact(added_storages.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB added storage capacity",
                source,
            })?;
        self.streams
            .try_reserve_exact(added_streams.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB added stream capacity",
                source,
            })?;
        self.storages.extend(added_storages);
        self.streams.extend(added_streams);
        Ok(())
    }

    pub(crate) fn remove_object(
        &mut self,
        path: &[String],
        limits: Limits,
    ) -> Result<(), OleError> {
        let found = self.storages.iter().any(|storage| storage.path() == path);
        if !found {
            return Err(OleError::InvalidFormat(format!(
                "object storage {path:?} not found"
            )));
        }
        self.storages
            .retain(|storage| !storage.path().starts_with(path));
        self.streams
            .retain(|stream| !stream.path().starts_with(path));
        self.check(limits)
    }

    pub(crate) fn render(&self) -> Result<Vec<u8>, OleError> {
        self.render_with_layout(None, SectorLayoutPolicy::default())
    }

    /// Renders the package, optionally reusing `source`'s sector layout.
    ///
    /// `source` is the artifact this package was captured from. Under
    /// [`SectorLayoutPolicy::Reuse`] the writer keeps that artifact's sector
    /// assignment wherever the package's stream and storage set still matches
    /// it, and falls back to the from-scratch serialization otherwise.
    pub(crate) fn render_with_layout(
        &self,
        source: Option<&[u8]>,
        policy: SectorLayoutPolicy,
    ) -> Result<Vec<u8>, OleError> {
        let mut writer = OleWriter::with_sector_size(self.sector_size)?;
        writer.set_sector_layout_policy(policy);
        if let Some(source) = source {
            writer.adopt_source_layout(source)?;
        }
        // A missing class ID is an explicit clear for a source-backed render;
        // otherwise the source-layout planner would preserve the old bytes and
        // make a directory transaction that removed a CLSID ineffective.
        writer.set_root_clsid(self.root_clsid.map_or([0; 16], |clsid| *clsid.as_bytes()));
        if let Some(root) = self
            .root_directory
            .filter(|root| root.kind() == EntryKind::Root)
        {
            // Only a captured SID-zero root may project its raw directory
            // fields onto the rendered root.  Promoted nested storages use
            // the writer's normative fresh-root defaults.
            writer.set_root_state_bits(root.state_bits());
            writer.set_root_modified_time(root.modified_time());
            writer.set_root_creation_time_from_source(root.creation_time());
        } else {
            // Supply the fresh-root defaults explicitly as well. The package's
            // directory metadata is then authoritative for every entry, so a
            // reused source layout can never publish source state or
            // timestamps that this package does not hold.
            writer.set_root_state_bits(0);
            writer.set_root_modified_time(0);
            writer.set_root_creation_time(0)?;
        }
        let mut storages = self.storages.clone();
        storages.sort_by(|left, right| {
            left.path()
                .len()
                .cmp(&right.path().len())
                .then_with(|| left.path().cmp(right.path()))
        });
        for storage in &storages {
            let refs = path_refs(storage.path());
            writer.create_storage(&refs)?;
            writer.set_storage_clsid(
                &refs,
                storage
                    .class_id()
                    .map_or([0; 16], |clsid| *clsid.as_bytes()),
            )?;
            let directory = storage.directory();
            if directory.state_bits() != 0
                || directory.creation_time() != 0
                || directory.modified_time() != 0
            {
                writer.set_storage_metadata(
                    &refs,
                    directory.state_bits(),
                    directory.creation_time(),
                    directory.modified_time(),
                )?;
            }
        }
        for stream in &self.streams {
            let refs = path_refs(stream.path());
            if let Some(directory) = stream.directory() {
                if directory.state_bits() != 0
                    || directory.creation_time() != 0
                    || directory.modified_time() != 0
                {
                    writer.create_stream_shared_with_metadata(
                        &refs,
                        stream.bytes_shared(),
                        directory.state_bits(),
                        directory.creation_time(),
                        directory.modified_time(),
                    )?;
                    continue;
                }
            }
            writer.create_stream_shared(&refs, stream.bytes_shared())?;
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output)?;
        Ok(output.into_inner())
    }

    /// Publishes equal-length stream edits through the validated source-backed
    /// overlay path.  The source model and directory topology must be the
    /// original one; a length change, entry change, or metadata edit returns
    /// `Ok(None)` so the caller can use the sector-layout writer instead.
    ///
    /// The composed positional view is reopened and every stream is read back
    /// before the returned bytes are materialized.  This keeps the 0617
    /// pre-emission invariant: a sink never observes an unvalidated candidate,
    /// and untouched streams are checked against their captured source bytes.
    ///
    /// `source` is the editor's immutable original allocation, so the CFB
    /// reader is opened with sealed ownership: the plan computes its source
    /// and target digests once, and neither the composed view nor the
    /// emission re-hashes an artifact that cannot have changed.
    pub(crate) fn render_copy_through(
        &self,
        baseline: &Self,
        source: &Arc<Vec<u8>>,
        limits: Limits,
    ) -> Result<Option<Vec<u8>>, OleError> {
        let Some(overlays) = self.same_length_overlays(baseline, limits)? else {
            return Ok(None);
        };
        if overlays.is_empty() {
            return Ok(Some(source.as_ref().clone()));
        }
        // The overlay plan changes payload spans only. A version-3 directory
        // entry stores a reserved high size word that the CFB writer
        // canonicalizes to zero even for an unchanged empty stream. If an
        // adopted source carries a nonzero reserved word, decline this path so
        // the source-layout writer can normalize it before publication.
        match source_v3_stream_size_needs_normalization(source) {
            Ok(true) => return Ok(None),
            Ok(false) => {},
            Err(error @ OleError::Allocation { .. }) => return Err(error),
            Err(_) => return Ok(None),
        }

        // The version only has to stay stable while this call's views are
        // alive; the allocation identity serves, as it does for the XLS
        // comments and visibility owners.
        let version = SourceVersion::new(source.as_ptr() as usize as u64, source.len() as u64);
        let shared = match SharedOleFile::open_owned_vec(Arc::clone(source), version) {
            Ok(shared) => shared,
            Err(error) => return overlay_fallback(error.into()),
        };
        let overlay_limits = match OverlayLimits::new(
            limits.max_streams.min(65_536),
            65_536,
            limits.max_total_size,
        ) {
            Ok(limits) => limits,
            Err(error) => return overlay_fallback(error),
        };
        let plan = match shared.plan_same_length_stream_overlays(overlays, overlay_limits) {
            Ok(plan) => plan,
            Err(error) => return overlay_fallback(error),
        };

        // Reopen the composed read-only view before allocating the output.
        // This validates the complete CFB partition and the stream identities
        // for both changed and untouched streams without a second artifact.
        let composed = match plan.composed_source() {
            Ok(composed) => composed,
            Err(error) => return overlay_fallback(error),
        };
        let candidate = match SharedOleFile::open(Arc::new(composed)) {
            Ok(candidate) => candidate,
            Err(error) => return overlay_fallback(error.into()),
        };
        for stream in &self.streams {
            let refs: Vec<&str> = stream.path().iter().map(String::as_str).collect();
            let bytes = match candidate.open_stream(&refs) {
                Ok(bytes) => bytes,
                Err(error) => return overlay_fallback(error.into()),
            };
            if bytes.as_slice() != stream.bytes() {
                return Ok(None);
            }
            if let Some(original) = baseline.stream(stream.path())
                && original == stream.bytes()
                && bytes.as_slice() != original
            {
                return Ok(None);
            }
        }

        let mut output = Vec::new();
        output
            .try_reserve_exact(source.len())
            .map_err(|source| OleError::Allocation {
                resource: "copy-through output",
                source,
            })?;
        if let Err(error) = plan.write_to(&mut output) {
            return overlay_fallback(error);
        }
        Ok(Some(output))
    }

    fn same_length_overlays(
        &self,
        baseline: &Self,
        limits: Limits,
    ) -> Result<Option<Vec<SameLengthStreamOverlay>>, OleError> {
        if self.streams.len() > limits.max_streams {
            return Ok(None);
        }
        // The overlay republishes the source directory image unchanged, so
        // any directory difference, including the root entry's state bits and
        // timestamps, declines to the layout writer.
        if self.sector_size != baseline.sector_size
            || self.root_clsid != baseline.root_clsid
            || self.root_directory != baseline.root_directory
            || self.storages != baseline.storages
            || self.streams.len() != baseline.streams.len()
        {
            return Ok(None);
        }
        let mut by_path = HashMap::new();
        by_path
            .try_reserve(baseline.streams.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB copy-through stream index",
                source,
            })?;
        for stream in &baseline.streams {
            by_path.insert(stream.path(), stream);
        }

        let mut overlays = Vec::new();
        overlays
            .try_reserve(self.streams.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB copy-through overlays",
                source,
            })?;
        for stream in &self.streams {
            let Some(original) = by_path.get(stream.path()).copied() else {
                return Ok(None);
            };
            if stream.bytes().len() != original.bytes().len() {
                return Ok(None);
            }
            if stream.directory() != original.directory() {
                return Ok(None);
            }
            if stream.bytes() == original.bytes() {
                continue;
            }
            overlays.push(SameLengthStreamOverlay::new(
                stream.path().to_vec(),
                stream.bytes_shared(),
            ));
        }
        if by_path.len() != self.streams.len() {
            return Ok(None);
        }
        Ok(Some(overlays))
    }

    pub(crate) fn check(&self, limits: Limits) -> Result<(), OleError> {
        limits.validate()?;
        check_counts(
            limits,
            self.storages.len(),
            self.streams.len(),
            stream_bytes_total(self.streams.iter())?,
        )
    }

    pub(crate) fn check_object_limits(&self, limits: Limits) -> Result<(), OleError> {
        limits.validate()?;
        if self.storages.len() > limits.max_object_storages()
            || self.streams.len() > limits.max_object_streams()
        {
            return Err(OleError::InvalidFormat(
                "selected object exceeds capture limits".into(),
            ));
        }
        Ok(())
    }

    fn object_from_root(
        &self,
        target: Target,
        storage: Storage,
        limits: Limits,
    ) -> Result<Object, OleError> {
        if self.storages.len() > limits.max_object_storages()
            || self.streams.len() > limits.max_object_streams()
        {
            return Err(OleError::InvalidFormat(
                "selected object exceeds capture limits".into(),
            ));
        }
        let compound = self.render()?;
        if compound.len() as u64 > limits.max_object_size {
            return Err(OleError::InvalidFormat(
                "selected object exceeds size limit".into(),
            ));
        }
        let storages = self.storages.clone();
        let streams = self.streams.clone();
        Ok(Object::new(
            target,
            storage,
            storages,
            streams,
            Arc::from(compound),
        ))
    }
}

fn overlay_fallback(error: OverlayError) -> Result<Option<Vec<u8>>, OleError> {
    match error {
        OverlayError::Unavailable { .. }
        | OverlayError::Ole(_)
        | OverlayError::SourceChanged { .. }
        | OverlayError::SourceFingerprintChanged { .. }
        | OverlayError::PreconditionFailed { .. }
        | OverlayError::TargetFingerprintChanged { .. } => Ok(None),
        OverlayError::Allocation { resource, source } => {
            Err(OleError::Allocation { resource, source })
        },
        OverlayError::Io(source) => Err(OleError::Io(source)),
        OverlayError::Committed { source } => Err(OleError::Committed { source }),
        OverlayError::IncompleteOutput { source, .. } => overlay_fallback(*source),
        _ => Ok(None),
    }
}

/// Returns whether a version-3 source has a nonzero reserved high size word
/// in any stream directory entry. The source has already passed the ordinary
/// CFB parser at this point, so this bounded raw walk only decides whether the
/// payload-only overlay may preserve the directory image. DIFAT sources are
/// conservatively declined to the layout writer as well.
fn source_v3_stream_size_needs_normalization(source: &[u8]) -> Result<bool, OleError> {
    const MAJOR_VERSION_OFFSET: usize = 0x1A;
    const SECTOR_SHIFT_OFFSET: usize = 0x1E;
    const FAT_COUNT_OFFSET: usize = 0x2C;
    const FIRST_DIRECTORY_SECTOR_OFFSET: usize = 0x30;
    const DIFAT_SECTOR_COUNT_OFFSET: usize = 0x48;
    const HEADER_DIFAT_OFFSET: usize = 0x4C;
    const DIRECTORY_ENTRY_SIZE: usize = 128;
    const ENTRY_TYPE_OFFSET: usize = 0x42;
    const ENTRY_SIZE_HIGH_OFFSET: usize = 0x7C;
    const ENDOFCHAIN: u32 = 0xFFFF_FFFE;
    const HEADER_DIFAT_ENTRIES: usize = 109;

    let major = u16::from_le_bytes(
        source
            .get(MAJOR_VERSION_OFFSET..MAJOR_VERSION_OFFSET + 2)
            .ok_or_else(|| OleError::InvalidFormat("CFB header is truncated".into()))?
            .try_into()
            .map_err(|_| OleError::InvalidFormat("CFB header version is truncated".into()))?,
    );
    if major != 3 {
        return Ok(false);
    }
    let shift = u16::from_le_bytes(
        source
            .get(SECTOR_SHIFT_OFFSET..SECTOR_SHIFT_OFFSET + 2)
            .ok_or_else(|| OleError::InvalidFormat("CFB sector shift is truncated".into()))?
            .try_into()
            .map_err(|_| OleError::InvalidFormat("CFB sector shift is truncated".into()))?,
    );
    let sector_size = match shift {
        9 => 512usize,
        12 => 4096usize,
        _ => {
            return Err(OleError::InvalidFormat(
                "CFB sector shift is invalid".into(),
            ));
        },
    };
    let header = source
        .get(..sector_size)
        .ok_or_else(|| OleError::InvalidFormat("CFB header is truncated".into()))?;
    let read_u32 = |offset: usize| -> Result<u32, OleError> {
        Ok(u32::from_le_bytes(
            header
                .get(offset..offset + 4)
                .ok_or_else(|| OleError::InvalidFormat("CFB header field is truncated".into()))?
                .try_into()
                .map_err(|_| OleError::InvalidFormat("CFB header field is truncated".into()))?,
        ))
    };
    if read_u32(DIFAT_SECTOR_COUNT_OFFSET)? != 0 {
        return Ok(true);
    }
    let fat_count = usize::try_from(read_u32(FAT_COUNT_OFFSET)?)
        .map_err(|_| OleError::InvalidFormat("CFB FAT count does not fit usize".into()))?;
    if fat_count > HEADER_DIFAT_ENTRIES {
        return Ok(true);
    }
    let entries_per_sector = sector_size / 4;
    let fat_entries = fat_count
        .checked_mul(entries_per_sector)
        .ok_or_else(|| OleError::InvalidFormat("CFB FAT count overflows usize".into()))?;
    let mut fat = Vec::new();
    fat.try_reserve_exact(fat_entries)
        .map_err(|source| OleError::Allocation {
            resource: "CFB v3 overlay FAT",
            source,
        })?;
    for index in 0..fat_count {
        let difat_offset = HEADER_DIFAT_OFFSET
            .checked_add(index.checked_mul(4).ok_or_else(|| {
                OleError::InvalidFormat("CFB DIFAT offset overflows usize".into())
            })?)
            .ok_or_else(|| OleError::InvalidFormat("CFB DIFAT offset overflows usize".into()))?;
        let sector = u32::from_le_bytes(
            header
                .get(difat_offset..difat_offset + 4)
                .ok_or_else(|| OleError::InvalidFormat("CFB DIFAT entry is truncated".into()))?
                .try_into()
                .map_err(|_| OleError::InvalidFormat("CFB DIFAT entry is truncated".into()))?,
        );
        let start = usize::try_from(sector)
            .ok()
            .and_then(|sector| sector.checked_add(1))
            .and_then(|sector| sector.checked_mul(sector_size))
            .ok_or_else(|| OleError::InvalidFormat("CFB FAT sector offset overflows".into()))?;
        let end = start
            .checked_add(sector_size)
            .ok_or_else(|| OleError::InvalidFormat("CFB FAT sector end overflows".into()))?;
        let fat_sector = source
            .get(start..end)
            .ok_or_else(|| OleError::InvalidFormat("CFB FAT sector is truncated".into()))?;
        for word in fat_sector.chunks_exact(4) {
            fat.push(u32::from_le_bytes(word.try_into().map_err(|_| {
                OleError::InvalidFormat("CFB FAT word is truncated".into())
            })?));
        }
    }

    let mut sector = read_u32(FIRST_DIRECTORY_SECTOR_OFFSET)?;
    let max_steps = source.len() / sector_size + 1;
    for _ in 0..=max_steps {
        if sector == ENDOFCHAIN {
            return Ok(false);
        }
        let sector_index = usize::try_from(sector).map_err(|_| {
            OleError::InvalidFormat("CFB directory sector does not fit usize".into())
        })?;
        let start = sector_index
            .checked_add(1)
            .and_then(|value| value.checked_mul(sector_size))
            .ok_or_else(|| {
                OleError::InvalidFormat("CFB directory sector offset overflows".into())
            })?;
        let end = start
            .checked_add(sector_size)
            .ok_or_else(|| OleError::InvalidFormat("CFB directory sector end overflows".into()))?;
        let directory_sector = source
            .get(start..end)
            .ok_or_else(|| OleError::InvalidFormat("CFB directory sector is truncated".into()))?;
        for entry in directory_sector.chunks_exact(DIRECTORY_ENTRY_SIZE) {
            if entry[ENTRY_TYPE_OFFSET] == 2
                && entry[ENTRY_SIZE_HIGH_OFFSET..ENTRY_SIZE_HIGH_OFFSET + 4]
                    .iter()
                    .any(|byte| *byte != 0)
            {
                return Ok(true);
            }
        }
        sector = *fat
            .get(sector_index)
            .ok_or_else(|| OleError::InvalidFormat("CFB directory chain leaves FAT".into()))?;
    }
    Err(OleError::InvalidFormat(
        "CFB directory chain exceeds the source bound".into(),
    ))
}

struct Budget {
    storages: usize,
    streams: usize,
    bytes: u64,
    max_storages: usize,
    max_streams: usize,
    max_bytes: u64,
}

impl Budget {
    fn new(max_storages: usize, max_streams: usize, max_bytes: u64) -> Self {
        Self {
            storages: 0,
            streams: 0,
            bytes: 0,
            max_storages,
            max_streams,
            max_bytes,
        }
    }

    fn max_directory_entries(&self) -> usize {
        self.max_storages.saturating_add(self.max_streams)
    }

    fn charge_storage(&mut self) -> Result<(), OleError> {
        if self.storages >= self.max_storages {
            return Err(OleError::InvalidFormat(
                "CFB storage count exceeds capture limit".into(),
            ));
        }
        self.storages += 1;
        Ok(())
    }

    fn charge(&mut self, size: u64) -> Result<(), OleError> {
        if self.streams >= self.max_streams {
            return Err(OleError::InvalidFormat(
                "CFB stream count exceeds capture limit".into(),
            ));
        }
        let total = self
            .bytes
            .checked_add(size)
            .ok_or_else(|| OleError::InvalidFormat("CFB capture size overflow".into()))?;
        if total > self.max_bytes {
            return Err(OleError::InvalidFormat(
                "CFB captured stream bytes exceed size limit".into(),
            ));
        }
        self.streams += 1;
        self.bytes = total;
        Ok(())
    }
}

fn capture_container<R: Read + Seek>(
    ole: &mut OleFile<R>,
    path: &[String],
    package: &mut Package,
    budget: &mut Budget,
    limits: Limits,
) -> Result<(), OleError> {
    let max_entries = budget.max_directory_entries();
    let refs = path_refs(path);
    ole.visit_directory_entries(&refs, max_entries, |ole, sid| {
        let (metadata, name, size) = {
            let entry = ole.directory_entry_by_sid(sid).ok_or_else(|| {
                OleError::CorruptedFile("directory visitor returned an unknown SID".into())
            })?;
            (directory::decode(entry)?, entry.name.clone(), entry.size)
        };
        let mut child = path.to_vec();
        child.push(name);
        match metadata.kind() {
            EntryKind::Storage => {
                if child.len() > limits.max_storage_depth {
                    return Err(OleError::InvalidFormat(
                        "CFB storage nesting limit exceeded".into(),
                    ));
                }
                budget.charge_storage()?;
                package
                    .storages
                    .try_reserve(1)
                    .map_err(|source| OleError::Allocation {
                        resource: "CFB storage capture",
                        source,
                    })?;
                package.storages.push(Storage::new(child.clone(), metadata));
                capture_container(ole, &child, package, budget, limits)?;
            },
            EntryKind::Stream => {
                if size > limits.max_stream_size {
                    return Err(OleError::InvalidFormat(format!(
                        "stream {child:?} exceeds size limit"
                    )));
                }
                budget.charge(size)?;
                let data = ole.open_stream(&path_refs(&child))?;
                if data.len() as u64 != size {
                    return Err(OleError::InvalidFormat(format!(
                        "stream {child:?} size changed during capture"
                    )));
                }
                package
                    .streams
                    .try_reserve(1)
                    .map_err(|source| OleError::Allocation {
                        resource: "CFB stream capture",
                        source,
                    })?;
                package
                    .streams
                    .push(Stream::new(child, Arc::<[u8]>::from(data), Some(metadata)));
            },
            EntryKind::Root => {},
        }
        Ok(())
    })
}

fn capture_subtree<R: Read + Seek>(
    ole: &mut OleFile<R>,
    absolute: &[String],
    relative: &[String],
    package: &mut Package,
    budget: &mut Budget,
    limits: Limits,
) -> Result<(), OleError> {
    if relative.len() > limits.max_storage_depth {
        return Err(OleError::InvalidFormat(
            "object storage nesting limit exceeded".into(),
        ));
    }
    let current = join(absolute, relative);
    let max_entries = budget.max_directory_entries();
    let refs = path_refs(&current);
    ole.visit_directory_entries(&refs, max_entries, |ole, sid| {
        let (metadata, name, size) = {
            let entry = ole.directory_entry_by_sid(sid).ok_or_else(|| {
                OleError::CorruptedFile("directory visitor returned an unknown SID".into())
            })?;
            (directory::decode(entry)?, entry.name.clone(), entry.size)
        };
        let mut child = relative.to_vec();
        child.push(name);
        match metadata.kind() {
            EntryKind::Storage => {
                if child.len() > limits.max_storage_depth {
                    return Err(OleError::InvalidFormat(
                        "object storage nesting limit exceeded".into(),
                    ));
                }
                budget.charge_storage()?;
                package
                    .storages
                    .try_reserve(1)
                    .map_err(|source| OleError::Allocation {
                        resource: "CFB storage capture",
                        source,
                    })?;
                package.storages.push(Storage::new(child.clone(), metadata));
                capture_subtree(ole, absolute, &child, package, budget, limits)?;
            },
            EntryKind::Stream => {
                if size > limits.max_stream_size {
                    return Err(OleError::InvalidFormat(
                        "object stream size exceeds limit".into(),
                    ));
                }
                budget.charge(size)?;
                let data = ole.open_stream(&path_refs(&join(absolute, &child)))?;
                if data.len() as u64 != size {
                    return Err(OleError::InvalidFormat(
                        "object stream size changed during capture".into(),
                    ));
                }
                package
                    .streams
                    .try_reserve(1)
                    .map_err(|source| OleError::Allocation {
                        resource: "CFB stream capture",
                        source,
                    })?;
                package
                    .streams
                    .push(Stream::new(child, Arc::<[u8]>::from(data), Some(metadata)));
            },
            EntryKind::Root => {},
        }
        Ok(())
    })
}

fn find_storage<R: Read + Seek>(
    ole: &mut OleFile<R>,
    path: &[String],
    max_entries: usize,
) -> Result<Storage, OleError> {
    let (name, parent) = path
        .split_last()
        .ok_or_else(|| OleError::InvalidFormat("object target path is empty".into()))?;
    let refs = path_refs(parent);
    let mut metadata = None;
    ole.visit_directory_entries(&refs, max_entries, |ole, sid| {
        let entry = ole.directory_entry_by_sid(sid).ok_or_else(|| {
            OleError::CorruptedFile("directory visitor returned an unknown SID".into())
        })?;
        if metadata.is_none() && entry.entry_type == EntryKind::Storage.raw() && entry.name == *name
        {
            metadata = Some(directory::decode(entry)?);
        }
        Ok::<(), OleError>(())
    })?;
    let metadata = metadata
        .ok_or_else(|| OleError::InvalidFormat(format!("object storage {path:?} not found")))?;
    if metadata.kind() != EntryKind::Storage {
        return Err(OleError::InvalidFormat(format!(
            "object target path {path:?} is not a storage"
        )));
    }
    Ok(Storage::new(path.to_vec(), metadata))
}

pub(crate) fn open<R: Read + Seek>(ole: &OleFile<R>, max_entries: usize) -> Result<(), OleError> {
    reject_protected_container(ole, "object editing", max_entries)
}

fn path_refs(path: &[String]) -> Vec<&str> {
    path.iter().map(String::as_str).collect()
}

fn join(left: &[String], right: &[String]) -> Vec<String> {
    left.iter().chain(right).cloned().collect()
}

fn clone_path(path: &[String]) -> Result<Vec<String>, OleError> {
    join_checked(&[], path)
}

fn join_checked(left: &[String], right: &[String]) -> Result<Vec<String>, OleError> {
    let count = left
        .len()
        .checked_add(right.len())
        .ok_or_else(|| OleError::InvalidFormat("CFB path component count overflow".into()))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(count)
        .map_err(|source| OleError::Allocation {
            resource: "CFB object path",
            source,
        })?;
    for component in left.iter().chain(right) {
        let mut value = String::new();
        value
            .try_reserve_exact(component.len())
            .map_err(|source| OleError::Allocation {
                resource: "CFB object path component",
                source,
            })?;
        value.push_str(component);
        output.push(value);
    }
    Ok(output)
}

fn stream_bytes_total<'a>(mut streams: impl Iterator<Item = &'a Stream>) -> Result<u64, OleError> {
    streams.try_fold(0u64, |total, stream| {
        total
            .checked_add(stream.bytes().len() as u64)
            .ok_or_else(|| OleError::InvalidFormat("CFB capture size overflow".into()))
    })
}

fn check_counts(
    limits: Limits,
    storage_count: usize,
    stream_count: usize,
    total_bytes: u64,
) -> Result<(), OleError> {
    if storage_count > limits.max_objects.saturating_mul(limits.max_storage_depth) {
        return Err(OleError::InvalidFormat(
            "CFB storage count exceeds object capture limit".into(),
        ));
    }
    if stream_count > limits.max_streams {
        return Err(OleError::InvalidFormat(
            "CFB stream count exceeds package capture limit".into(),
        ));
    }
    if total_bytes > limits.max_total_size {
        return Err(OleError::InvalidFormat(
            "CFB captured stream bytes exceed total size limit".into(),
        ));
    }
    Ok(())
}
