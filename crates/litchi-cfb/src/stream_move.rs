//! Exact directory-only moves of existing CFB streams.

use crate::SharedOleFile;
use crate::consts::{DIRENTRY_SIZE, NOSTREAM, STGTY_ROOT, STGTY_STORAGE, STGTY_STREAM};
use crate::directory_name::{DirectoryNameData, directory_name_data};
use crate::overlay::{
    ArtifactFingerprint, OverlayError, OverlayLimits, PhysicalSpan, PublishReport, SourceSnapshot,
    ValidatedOverlayPlan, collect_chain_exact, finish_overlay_plan, path_refs, sector_offset,
    unavailable, validate_and_coalesce_spans,
};
use litchi_core::ReadAt;
use std::cmp::Ordering;
use std::io::Write;
use std::path::Path;
use std::sync::Arc;

/// One exact-source move or rename of an existing stream.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ExistingStreamMove {
    source: Vec<String>,
    destination: Vec<String>,
}

impl ExistingStreamMove {
    /// Creates a stream move. Both paths include the stream name.
    #[must_use]
    pub fn new(source: Vec<String>, destination: Vec<String>) -> Self {
        Self {
            source,
            destination,
        }
    }

    /// Exact source path resolved while planning.
    #[must_use]
    pub fn source(&self) -> &[String] {
        &self.source
    }

    /// Destination parent and new stream name.
    #[must_use]
    pub fn destination(&self) -> &[String] {
        &self.destination
    }
}

/// Finite input and derived-work bounds for one stream-move batch.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct StreamMoveLimits {
    max_moves: usize,
    max_path_components: usize,
    max_path_bytes: usize,
    max_directory_entries: usize,
    max_directory_bytes: usize,
    max_changed_spans: usize,
}

impl StreamMoveLimits {
    /// Creates a non-zero bounded stream-move policy.
    pub fn new(
        max_moves: usize,
        max_path_components: usize,
        max_path_bytes: usize,
        max_directory_entries: usize,
        max_directory_bytes: usize,
        max_changed_spans: usize,
    ) -> Result<Self, OverlayError> {
        if max_moves == 0
            || max_path_components == 0
            || max_path_bytes == 0
            || max_directory_entries == 0
            || max_directory_bytes == 0
            || max_changed_spans == 0
        {
            return Err(unavailable("stream move limits must be non-zero"));
        }
        Ok(Self {
            max_moves,
            max_path_components,
            max_path_bytes,
            max_directory_entries,
            max_directory_bytes,
            max_changed_spans,
        })
    }

    /// Maximum moves in one atomic batch.
    #[must_use]
    pub const fn max_moves(self) -> usize {
        self.max_moves
    }

    /// Maximum aggregate number of source and destination path components.
    #[must_use]
    pub const fn max_path_components(self) -> usize {
        self.max_path_components
    }

    /// Maximum aggregate UTF-8 bytes in source and destination paths.
    #[must_use]
    pub const fn max_path_bytes(self) -> usize {
        self.max_path_bytes
    }

    /// Maximum directory slots inspected by one batch.
    #[must_use]
    pub const fn max_directory_entries(self) -> usize {
        self.max_directory_entries
    }

    /// Maximum directory-stream bytes materialized by one batch.
    #[must_use]
    pub const fn max_directory_bytes(self) -> usize {
        self.max_directory_bytes
    }

    /// Maximum changed physical directory-sector fragments retained per plan.
    #[must_use]
    pub const fn max_changed_spans(self) -> usize {
        self.max_changed_spans
    }
}

impl Default for StreamMoveLimits {
    fn default() -> Self {
        Self {
            max_moves: 64,
            max_path_components: 1_024,
            max_path_bytes: 256 * 1024,
            max_directory_entries: 1_048_576,
            max_directory_bytes: 64 * 1024 * 1024,
            max_changed_spans: 131_072,
        }
    }
}

/// Validated exact forward and inverse plans for one atomic stream-move batch.
///
/// Only the CFB directory stream is materialized. Stream payload sectors,
/// FAT/MiniFAT allocation, slack, and every unrelated physical byte are copied
/// from the positional source unchanged during publication. This generic CFB
/// substrate is dependency-agnostic: format owners remain responsible for
/// rejecting moves that violate their semantic, signature, encryption, or DRM
/// policy.
///
/// The inverse is prepared during planning and restores the exact original
/// directory bytes, including the original sibling-tree shape. Plans are
/// source-backed rather than serialized; both directions retain complete
/// artifact fingerprints and source-version checks.
pub struct ValidatedStreamMovePlan {
    forward: ValidatedOverlayPlan,
    inverse: ValidatedOverlayPlan,
}

impl std::fmt::Debug for ValidatedStreamMovePlan {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ValidatedStreamMovePlan")
            .field("source_fingerprint", &self.source_fingerprint())
            .field("target_fingerprint", &self.target_fingerprint())
            .finish_non_exhaustive()
    }
}

impl ValidatedStreamMovePlan {
    /// Exact source artifact fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> ArtifactFingerprint {
        self.forward.source_fingerprint()
    }

    /// Exact moved artifact fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> ArtifactFingerprint {
        self.forward.target_fingerprint()
    }

    /// Checked forward directory-overlay plan.
    #[must_use]
    pub const fn forward(&self) -> &ValidatedOverlayPlan {
        &self.forward
    }

    /// Checked inverse plan whose source is the exact composed forward target.
    #[must_use]
    pub const fn inverse(&self) -> &ValidatedOverlayPlan {
        &self.inverse
    }

    /// Streams the complete moved artifact to a sequential sink.
    pub fn write_to<W: Write>(&self, writer: &mut W) -> Result<PublishReport, OverlayError> {
        self.forward.write_to(writer)
    }

    /// Atomically publishes the complete moved artifact to `path`.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<PublishReport, OverlayError> {
        self.forward.save(path)
    }
}

#[derive(Clone)]
struct PlannedMove {
    destination: Vec<String>,
    sid: u32,
    new_parent: u32,
    new_name: DirectoryNameData,
    changed: bool,
}

#[derive(Clone)]
struct Node {
    name: DirectoryNameData,
    left: u32,
    right: u32,
    child: u32,
    color: u8,
}

impl SharedOleFile {
    /// Plans a bounded atomic batch of existing-stream renames or moves.
    ///
    /// Source paths and destination parents are resolved against the same
    /// validated snapshot. Storage moves, implicit storage creation, and
    /// case-insensitive sibling collisions are refused. Final sibling sets are
    /// evaluated together, so swaps and cycles within a batch are well-defined.
    pub fn plan_stream_moves(
        &self,
        moves: Vec<ExistingStreamMove>,
        limits: StreamMoveLimits,
    ) -> Result<ValidatedStreamMovePlan, OverlayError> {
        self.check_source_version()?;
        validate_request_bounds(&moves, limits)?;
        let count = self.index.dir_entries.len();
        if count > limits.max_directory_entries {
            return Err(unavailable(format!(
                "directory entry count {count} exceeds limit {}",
                limits.max_directory_entries
            )));
        }
        let directory_bytes = count
            .checked_mul(DIRENTRY_SIZE)
            .ok_or_else(|| unavailable("directory byte length overflows usize"))?;
        if directory_bytes > limits.max_directory_bytes {
            return Err(unavailable(format!(
                "directory byte length {directory_bytes} exceeds limit {}",
                limits.max_directory_bytes
            )));
        }

        let source = SourceSnapshot {
            source: Arc::clone(&self.source),
            version: self.expected_version,
            length: self.index.file_size,
        };
        source.ensure_length()?;
        let sectors = directory_sectors(self, directory_bytes)?;
        let original = read_directory(&source, &sectors, self.index.sector_size, directory_bytes)?;
        let parents = directory_parents(self)?;
        let planned = resolve_moves(self, moves, &parents)?;
        let mut target = Vec::new();
        target
            .try_reserve_exact(original.len())
            .map_err(|source| OverlayError::Allocation {
                resource: "CFB target directory stream",
                source,
            })?;
        target.extend_from_slice(&original);
        rewrite_directory(self, &planned, &parents, &mut target)?;

        let overlay_limits = OverlayLimits::new(
            limits.max_moves,
            limits.max_changed_spans,
            u64::try_from(limits.max_directory_bytes)
                .map_err(|_| unavailable("directory byte limit does not fit u64"))?,
        )?;
        let forward_spans = directory_spans(
            &source,
            &sectors,
            self.index.sector_size,
            Arc::from(target),
            overlay_limits,
        )?;
        let verify_moves = planned.clone();
        let verify_index = Arc::clone(&self.index);
        let forward = finish_overlay_plan(source, forward_spans, move |_source, candidate| {
            verify_candidate(candidate, &verify_index, &verify_moves)
        })?;

        let composed = forward.composed_source()?;
        let inverse_source: Arc<dyn ReadAt> = Arc::new(composed);
        let inverse_view = SharedOleFile::open(Arc::clone(&inverse_source))?;
        let inverse_snapshot = SourceSnapshot {
            source: inverse_source,
            version: inverse_view.expected_version,
            length: inverse_view.index.file_size,
        };
        let inverse_spans = directory_spans(
            &inverse_snapshot,
            &sectors,
            self.index.sector_size,
            Arc::from(original),
            overlay_limits,
        )?;
        let original_index = Arc::clone(&self.index);
        let inverse = finish_overlay_plan(
            inverse_snapshot,
            inverse_spans,
            move |_source, candidate| verify_exact_metadata(candidate, &original_index),
        )?;
        if inverse.source_fingerprint() != forward.target_fingerprint()
            || inverse.target_fingerprint() != forward.source_fingerprint()
        {
            return Err(unavailable(
                "stream move inverse did not restore exact source bytes",
            ));
        }
        Ok(ValidatedStreamMovePlan { forward, inverse })
    }
}

fn validate_request_bounds(
    moves: &[ExistingStreamMove],
    limits: StreamMoveLimits,
) -> Result<(), OverlayError> {
    if moves.len() > limits.max_moves {
        return Err(unavailable(format!(
            "stream move count {} exceeds limit {}",
            moves.len(),
            limits.max_moves
        )));
    }
    let mut components = 0usize;
    let mut bytes = 0usize;
    for request in moves {
        if request.source.is_empty()
            || request.destination.is_empty()
            || request.source.iter().any(String::is_empty)
            || request.destination.iter().any(String::is_empty)
        {
            return Err(unavailable(
                "stream move paths must contain non-empty names",
            ));
        }
        for path in [&request.source, &request.destination] {
            components = components
                .checked_add(path.len())
                .ok_or_else(|| unavailable("aggregate path component count overflows usize"))?;
            for component in path {
                bytes = bytes
                    .checked_add(component.len())
                    .ok_or_else(|| unavailable("aggregate path bytes overflow usize"))?;
            }
        }
    }
    if components > limits.max_path_components {
        return Err(unavailable(format!(
            "aggregate path component count {components} exceeds limit {}",
            limits.max_path_components
        )));
    }
    if bytes > limits.max_path_bytes {
        return Err(unavailable(format!(
            "aggregate path bytes {bytes} exceed limit {}",
            limits.max_path_bytes
        )));
    }
    Ok(())
}

fn resolve_moves(
    file: &SharedOleFile,
    moves: Vec<ExistingStreamMove>,
    parents: &[u32],
) -> Result<Vec<PlannedMove>, OverlayError> {
    let mut planned = Vec::new();
    planned
        .try_reserve_exact(moves.len())
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB planned stream moves",
            source,
        })?;
    for request in moves {
        let source_refs = path_refs(&request.source)?;
        let entry = file.find_entry(&source_refs)?;
        if entry.entry_type != STGTY_STREAM {
            return Err(unavailable(format!(
                "stream move source {:?} is not a stream",
                request.source
            )));
        }
        if planned
            .iter()
            .any(|item: &PlannedMove| item.sid == entry.sid)
        {
            return Err(unavailable(format!(
                "stream move batch selects source {:?} more than once",
                request.source
            )));
        }
        let destination_parent = &request.destination[..request.destination.len() - 1];
        let new_parent = if destination_parent.is_empty() {
            0
        } else {
            let refs = path_refs(destination_parent)?;
            let parent = file.find_entry(&refs)?;
            if parent.entry_type != STGTY_STORAGE {
                return Err(unavailable(format!(
                    "stream move destination parent {:?} is not a storage",
                    destination_parent
                )));
            }
            parent.sid
        };
        let destination_name = request
            .destination
            .last()
            .ok_or_else(|| unavailable("stream move destination path is empty"))?;
        let new_name = directory_name_data(destination_name)
            .map_err(|error| unavailable(error.to_string()))?;
        let old_parent = *parents
            .get(entry.sid as usize)
            .ok_or_else(|| unavailable("stream SID does not fit directory parent map"))?;
        let old_name = file
            .index
            .dir_name_data
            .get(entry.sid as usize)
            .and_then(Option::as_ref)
            .ok_or_else(|| unavailable("stream SID has no validated name"))?;
        planned.push(PlannedMove {
            destination: request.destination,
            sid: entry.sid,
            new_parent,
            changed: old_parent != new_parent || old_name.utf16 != new_name.utf16,
            new_name,
        });
    }
    Ok(planned)
}

fn directory_parents(file: &SharedOleFile) -> Result<Vec<u32>, OverlayError> {
    let count = file.index.dir_entries.len();
    let mut parents = Vec::new();
    parents
        .try_reserve_exact(count)
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB directory parent map",
            source,
        })?;
    parents.resize(count, NOSTREAM);
    let mut stack = Vec::new();
    stack
        .try_reserve_exact(count)
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB directory parent traversal",
            source,
        })?;
    for owner in file
        .index
        .dir_entries
        .iter()
        .flatten()
        .filter(|entry| entry.entry_type == STGTY_ROOT || entry.entry_type == STGTY_STORAGE)
    {
        if owner.sid_child != NOSTREAM {
            stack.push(owner.sid_child);
        }
        while let Some(sid) = stack.pop() {
            let entry = file
                .index
                .dir_entries
                .get(sid as usize)
                .and_then(Option::as_ref)
                .ok_or_else(|| unavailable("validated directory contains an invalid child SID"))?;
            let slot = parents
                .get_mut(sid as usize)
                .ok_or_else(|| unavailable("directory child SID does not fit parent map"))?;
            if *slot != NOSTREAM {
                return Err(unavailable("directory entry has multiple storage parents"));
            }
            *slot = owner.sid;
            if entry.sid_left != NOSTREAM {
                stack.push(entry.sid_left);
            }
            if entry.sid_right != NOSTREAM {
                stack.push(entry.sid_right);
            }
        }
    }
    Ok(parents)
}

fn rewrite_directory(
    file: &SharedOleFile,
    planned: &[PlannedMove],
    original_parents: &[u32],
    bytes: &mut [u8],
) -> Result<(), OverlayError> {
    let count = file.index.dir_entries.len();
    let mut parents = Vec::new();
    parents
        .try_reserve_exact(original_parents.len())
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB updated directory parent map",
            source,
        })?;
    parents.extend_from_slice(original_parents);
    let mut nodes = Vec::new();
    nodes
        .try_reserve_exact(count)
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB mutable directory nodes",
            source,
        })?;
    for index in 0..count {
        let entry = file.index.dir_entries[index].as_ref();
        let name = file.index.dir_name_data[index].clone();
        nodes.push(match (entry, name) {
            (Some(entry), Some(name)) => Some(Node {
                name,
                left: entry.sid_left,
                right: entry.sid_right,
                child: entry.sid_child,
                color: bytes[index * DIRENTRY_SIZE + 67],
            }),
            (None, None) => None,
            _ => return Err(unavailable("directory metadata caches disagree")),
        });
    }
    for item in planned {
        if item.changed {
            parents[item.sid as usize] = item.new_parent;
            nodes[item.sid as usize]
                .as_mut()
                .ok_or_else(|| unavailable("moved stream SID is empty"))?
                .name = item.new_name.clone();
        }
    }

    let mut affected = Vec::new();
    affected
        .try_reserve_exact(count)
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB affected directory owners",
            source,
        })?;
    affected.resize(count, false);
    for item in planned {
        if !item.changed {
            continue;
        }
        let old_parent = *original_parents
            .get(item.sid as usize)
            .ok_or_else(|| unavailable("stream SID does not fit directory parent map"))?;
        *affected
            .get_mut(old_parent as usize)
            .ok_or_else(|| unavailable("old parent SID does not fit directory"))? = true;
        *affected
            .get_mut(item.new_parent as usize)
            .ok_or_else(|| unavailable("new parent SID does not fit directory"))? = true;
    }

    let mut child_counts = Vec::new();
    child_counts
        .try_reserve_exact(count)
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB directory child counts",
            source,
        })?;
    child_counts.resize(count, 0usize);
    for (sid, parent) in parents.iter().copied().enumerate().skip(1) {
        if nodes[sid].is_some() {
            let slot = child_counts
                .get_mut(parent as usize)
                .ok_or_else(|| unavailable("directory parent SID does not fit child counts"))?;
            *slot = slot
                .checked_add(1)
                .ok_or_else(|| unavailable("directory child count overflows usize"))?;
        }
    }
    let mut children = Vec::new();
    children
        .try_reserve_exact(count)
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB directory child lists",
            source,
        })?;
    for child_count in child_counts {
        let mut list = Vec::new();
        list.try_reserve_exact(child_count)
            .map_err(|source| OverlayError::Allocation {
                resource: "CFB directory child list",
                source,
            })?;
        children.push(list);
    }
    for (sid, parent) in parents.iter().copied().enumerate().skip(1) {
        if nodes[sid].is_some() {
            children[parent as usize].push(sid as u32);
        }
    }
    let mut tree_parents = Vec::new();
    tree_parents
        .try_reserve_exact(count)
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB sibling-tree parent scratch",
            source,
        })?;
    tree_parents.resize(count, NOSTREAM);
    for (owner, is_affected) in affected.into_iter().enumerate() {
        if is_affected {
            tree_parents.fill(NOSTREAM);
            link_children(
                owner as u32,
                &mut children[owner],
                &mut nodes,
                &mut tree_parents,
            )?;
        }
    }

    for item in planned {
        if item.changed {
            encode_name(bytes, item.sid as usize, &item.new_name)?;
        }
    }
    for (sid, node) in nodes.iter().enumerate() {
        let Some(node) = node else { continue };
        let base = sid
            .checked_mul(DIRENTRY_SIZE)
            .ok_or_else(|| unavailable("directory slot offset overflows usize"))?;
        bytes[base + 67] = node.color;
        bytes[base + 68..base + 72].copy_from_slice(&node.left.to_le_bytes());
        bytes[base + 72..base + 76].copy_from_slice(&node.right.to_le_bytes());
        bytes[base + 76..base + 80].copy_from_slice(&node.child.to_le_bytes());
    }
    Ok(())
}

fn encode_name(bytes: &mut [u8], sid: usize, name: &DirectoryNameData) -> Result<(), OverlayError> {
    let base = sid
        .checked_mul(DIRENTRY_SIZE)
        .ok_or_else(|| unavailable("directory name slot offset overflows usize"))?;
    bytes[base..base + 64].fill(0);
    for (index, unit) in name.utf16.iter().copied().enumerate() {
        let offset = base + index * 2;
        bytes[offset..offset + 2].copy_from_slice(&unit.to_le_bytes());
    }
    let byte_len = u16::try_from((name.utf16.len() + 1) * 2)
        .map_err(|_| unavailable("directory name byte length does not fit u16"))?;
    bytes[base + 64..base + 66].copy_from_slice(&byte_len.to_le_bytes());
    Ok(())
}

fn compare_nodes(left: u32, right: u32, nodes: &[Option<Node>]) -> Ordering {
    nodes[left as usize]
        .as_ref()
        .expect("active child")
        .name
        .compare(&nodes[right as usize].as_ref().expect("active child").name)
}

fn color(nodes: &[Option<Node>], sid: u32) -> u8 {
    if sid == NOSTREAM {
        1
    } else {
        nodes[sid as usize].as_ref().expect("active node").color
    }
}

fn link_children(
    owner: u32,
    children: &mut [u32],
    nodes: &mut [Option<Node>],
    parents: &mut [u32],
) -> Result<(), OverlayError> {
    children.sort_unstable_by(|left, right| compare_nodes(*left, *right, nodes));
    for pair in children.windows(2) {
        if compare_nodes(pair[0], pair[1], nodes) == Ordering::Equal {
            return Err(unavailable(format!(
                "stream move creates a case-insensitive sibling collision at SID {}",
                pair[1]
            )));
        }
    }
    if children.is_empty() {
        nodes[owner as usize].as_mut().expect("active owner").child = NOSTREAM;
        return Ok(());
    }
    let mut root = NOSTREAM;
    for &sid in children.iter() {
        let node = nodes[sid as usize].as_mut().expect("active child");
        node.left = NOSTREAM;
        node.right = NOSTREAM;
        node.color = 0;
        let mut parent = NOSTREAM;
        let mut cursor = root;
        let mut ordering = Ordering::Equal;
        while cursor != NOSTREAM {
            parent = cursor;
            ordering = compare_nodes(sid, cursor, nodes);
            cursor = if ordering == Ordering::Less {
                nodes[cursor as usize].as_ref().expect("active node").left
            } else {
                nodes[cursor as usize].as_ref().expect("active node").right
            };
        }
        parents[sid as usize] = parent;
        if parent == NOSTREAM {
            root = sid;
        } else if ordering == Ordering::Less {
            nodes[parent as usize].as_mut().expect("active parent").left = sid;
        } else {
            nodes[parent as usize]
                .as_mut()
                .expect("active parent")
                .right = sid;
        }
        fix_after_insert(&mut root, sid, parents, nodes);
    }
    nodes[owner as usize].as_mut().expect("active owner").child = root;
    Ok(())
}

fn rotate_left(root: &mut u32, pivot: u32, parents: &mut [u32], nodes: &mut [Option<Node>]) {
    let child = nodes[pivot as usize].as_ref().expect("active pivot").right;
    let child_left = nodes[child as usize].as_ref().expect("active child").left;
    nodes[pivot as usize].as_mut().expect("active pivot").right = child_left;
    if child_left != NOSTREAM {
        parents[child_left as usize] = pivot;
    }
    let pivot_parent = parents[pivot as usize];
    parents[child as usize] = pivot_parent;
    if pivot_parent == NOSTREAM {
        *root = child;
    } else if pivot
        == nodes[pivot_parent as usize]
            .as_ref()
            .expect("active parent")
            .left
    {
        nodes[pivot_parent as usize]
            .as_mut()
            .expect("active parent")
            .left = child;
    } else {
        nodes[pivot_parent as usize]
            .as_mut()
            .expect("active parent")
            .right = child;
    }
    nodes[child as usize].as_mut().expect("active child").left = pivot;
    parents[pivot as usize] = child;
}

fn rotate_right(root: &mut u32, pivot: u32, parents: &mut [u32], nodes: &mut [Option<Node>]) {
    let child = nodes[pivot as usize].as_ref().expect("active pivot").left;
    let child_right = nodes[child as usize].as_ref().expect("active child").right;
    nodes[pivot as usize].as_mut().expect("active pivot").left = child_right;
    if child_right != NOSTREAM {
        parents[child_right as usize] = pivot;
    }
    let pivot_parent = parents[pivot as usize];
    parents[child as usize] = pivot_parent;
    if pivot_parent == NOSTREAM {
        *root = child;
    } else if pivot
        == nodes[pivot_parent as usize]
            .as_ref()
            .expect("active parent")
            .right
    {
        nodes[pivot_parent as usize]
            .as_mut()
            .expect("active parent")
            .right = child;
    } else {
        nodes[pivot_parent as usize]
            .as_mut()
            .expect("active parent")
            .left = child;
    }
    nodes[child as usize].as_mut().expect("active child").right = pivot;
    parents[pivot as usize] = child;
}

fn fix_after_insert(root: &mut u32, mut sid: u32, parents: &mut [u32], nodes: &mut [Option<Node>]) {
    while color(nodes, parents[sid as usize]) == 0 {
        let parent = parents[sid as usize];
        let grandparent = parents[parent as usize];
        if parent
            == nodes[grandparent as usize]
                .as_ref()
                .expect("active grandparent")
                .left
        {
            let uncle = nodes[grandparent as usize]
                .as_ref()
                .expect("active grandparent")
                .right;
            if color(nodes, uncle) == 0 {
                nodes[parent as usize]
                    .as_mut()
                    .expect("active parent")
                    .color = 1;
                nodes[uncle as usize].as_mut().expect("active uncle").color = 1;
                nodes[grandparent as usize]
                    .as_mut()
                    .expect("active grandparent")
                    .color = 0;
                sid = grandparent;
            } else {
                if sid
                    == nodes[parent as usize]
                        .as_ref()
                        .expect("active parent")
                        .right
                {
                    sid = parent;
                    rotate_left(root, sid, parents, nodes);
                }
                let rotated_parent = parents[sid as usize];
                let rotated_grandparent = parents[rotated_parent as usize];
                nodes[rotated_parent as usize]
                    .as_mut()
                    .expect("active parent")
                    .color = 1;
                nodes[rotated_grandparent as usize]
                    .as_mut()
                    .expect("active grandparent")
                    .color = 0;
                rotate_right(root, rotated_grandparent, parents, nodes);
            }
        } else {
            let uncle = nodes[grandparent as usize]
                .as_ref()
                .expect("active grandparent")
                .left;
            if color(nodes, uncle) == 0 {
                nodes[parent as usize]
                    .as_mut()
                    .expect("active parent")
                    .color = 1;
                nodes[uncle as usize].as_mut().expect("active uncle").color = 1;
                nodes[grandparent as usize]
                    .as_mut()
                    .expect("active grandparent")
                    .color = 0;
                sid = grandparent;
            } else {
                if sid == nodes[parent as usize].as_ref().expect("active parent").left {
                    sid = parent;
                    rotate_right(root, sid, parents, nodes);
                }
                let rotated_parent = parents[sid as usize];
                let rotated_grandparent = parents[rotated_parent as usize];
                nodes[rotated_parent as usize]
                    .as_mut()
                    .expect("active parent")
                    .color = 1;
                nodes[rotated_grandparent as usize]
                    .as_mut()
                    .expect("active grandparent")
                    .color = 0;
                rotate_left(root, rotated_grandparent, parents, nodes);
            }
        }
    }
    nodes[*root as usize].as_mut().expect("active root").color = 1;
}

fn directory_sectors(
    file: &SharedOleFile,
    directory_bytes: usize,
) -> Result<Vec<u32>, OverlayError> {
    let sector_count = directory_bytes
        .checked_div(file.index.sector_size)
        .ok_or_else(|| unavailable("directory sector size is zero"))?;
    collect_chain_exact(
        &file.index.fat,
        file.index.first_dir_sector,
        sector_count,
        "directory",
    )
}

fn read_directory(
    source: &SourceSnapshot,
    sectors: &[u32],
    sector_size: usize,
    byte_len: usize,
) -> Result<Vec<u8>, OverlayError> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(byte_len)
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB directory stream snapshot",
            source,
        })?;
    bytes.resize(byte_len, 0);
    for (ordinal, sector) in sectors.iter().copied().enumerate() {
        let physical = sector_offset(sector, sector_size)?;
        let logical = ordinal * sector_size;
        let present = usize::try_from(
            source
                .length
                .saturating_sub(physical)
                .min(sector_size as u64),
        )
        .map_err(|_| unavailable("directory sector length does not fit usize"))?;
        source.read_exact(physical, &mut bytes[logical..logical + present])?;
    }
    Ok(bytes)
}

fn directory_spans(
    source: &SourceSnapshot,
    sectors: &[u32],
    sector_size: usize,
    replacement: Arc<[u8]>,
    limits: OverlayLimits,
) -> Result<Vec<PhysicalSpan>, OverlayError> {
    let mut spans = Vec::new();
    spans
        .try_reserve_exact(sectors.len().min(limits.max_spans()))
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB directory overlay spans",
            source,
        })?;
    let mut comparison = Vec::new();
    comparison
        .try_reserve_exact(sector_size)
        .map_err(|source| OverlayError::Allocation {
            resource: "CFB directory sector comparison",
            source,
        })?;
    comparison.resize(sector_size, 0u8);
    for (ordinal, sector) in sectors.iter().copied().enumerate() {
        let physical = sector_offset(sector, sector_size)?;
        let logical = ordinal * sector_size;
        let present = usize::try_from(
            source
                .length
                .saturating_sub(physical)
                .min(sector_size as u64),
        )
        .map_err(|_| unavailable("directory sector length does not fit usize"))?;
        source.read_exact(physical, &mut comparison[..present])?;
        if comparison[..present] != replacement[logical..logical + present] {
            if spans.len() == limits.max_spans() {
                return Err(unavailable(format!(
                    "changed directory fragment count exceeds limit {}",
                    limits.max_spans()
                )));
            }
            spans
                .try_reserve(1)
                .map_err(|source| OverlayError::Allocation {
                    resource: "CFB directory overlay spans",
                    source,
                })?;
            spans.push(PhysicalSpan {
                offset: physical,
                replacement: Arc::clone(&replacement),
                replacement_range: logical..logical + present,
            });
        }
    }
    spans.sort_unstable_by_key(|span| span.offset);
    validate_and_coalesce_spans(source, limits, &mut spans)?;
    Ok(spans)
}

fn verify_candidate(
    candidate: &SharedOleFile,
    source_index: &crate::file::ParsedOleIndex,
    moves: &[PlannedMove],
) -> Result<(), OverlayError> {
    verify_exact_metadata(candidate, source_index)?;
    for item in moves {
        let refs = path_refs(&item.destination)?;
        let entry = candidate.find_entry(&refs)?;
        if entry.sid != item.sid || entry.entry_type != STGTY_STREAM {
            return Err(unavailable(format!(
                "composed destination {:?} does not identify source SID {}",
                item.destination, item.sid
            )));
        }
    }
    Ok(())
}

fn verify_exact_metadata(
    candidate: &SharedOleFile,
    source: &crate::file::ParsedOleIndex,
) -> Result<(), OverlayError> {
    if candidate.index.dir_entries.len() != source.dir_entries.len()
        || candidate.index.fat != source.fat
        || candidate.index.minifat != source.minifat
    {
        return Err(unavailable("stream move changed CFB allocation metadata"));
    }
    for (before, after) in source.dir_entries.iter().zip(&candidate.index.dir_entries) {
        match (before, after) {
            (None, None) => {},
            (Some(before), Some(after))
                if before.sid == after.sid
                    && before.entry_type == after.entry_type
                    && before.clsid == after.clsid
                    && before.start_sector == after.start_sector
                    && before.size == after.size
                    && before.is_minifat == after.is_minifat => {},
            _ => {
                return Err(unavailable(
                    "stream move changed non-topological directory metadata",
                ));
            },
        }
    }
    Ok(())
}
