//! Unified title and linked-section transactions.

use litchi_core::{
    BlobBundle, BlobId, BlobLimits, ConflictSet, Error, ForwardOnly, History as CoreHistory,
    HistoryLimits, Patch as CorePatch, PatchLimits, PatchOperation, Position, Result, Reversible,
    ReversibleOperation,
};
use litchi_odf_common::core::{MetaXmlPatch, metadata::Metadata as OdfMetadata, patch_meta_xml};
use serde_json::Value;
use std::{
    collections::{BTreeMap, HashMap},
    sync::Arc,
};

use crate::{Master, link::Selector};

const MAX_PACKAGE_BYTES: usize = 256 * 1024 * 1024;
const MAX_WIRE_JSON_BYTES: usize = 768 * 1024 * 1024;
const DURABLE_FORMAT: &str = "litchi.odm";
const DURABLE_OPERATION: &str = "package.replace";
const DURABLE_TARGET: &str = "master";
const SOURCE_PRECONDITION: &str = "source_sha256";

/// One atomic edit derived from an immutable master snapshot.
pub struct Edit<'source> {
    source: &'source Master,
    title_before: Option<String>,
    title_after: Option<String>,
    links: BTreeMap<usize, String>,
}

impl<'source> Edit<'source> {
    pub(crate) fn new(source: &'source Master) -> Self {
        let title = source.title().map(str::to_owned);
        Self {
            source,
            title_before: title.clone(),
            title_after: title,
            links: BTreeMap::new(),
        }
    }

    /// Returns the staged title.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title_after.as_deref()
    }

    /// Stages a bounded XML 1.0 title.
    ///
    /// # Errors
    ///
    /// Returns an error when the title exceeds the limit or contains a
    /// character forbidden by XML 1.0.
    pub fn set_title(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        let title = value.into();
        crate::title::validate_title(&title)?;
        self.title_after = Some(title);
        Ok(self)
    }

    /// Stages removal of the title element.
    pub fn clear_title(&mut self) -> &mut Self {
        self.title_after = None;
        self
    }

    /// Stages one existing linked-section target by exact section name or
    /// checked semantic position.
    ///
    /// Targets remain inert and are never resolved, opened, or fetched.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing selector or invalid XML target.
    pub fn set_link<'selector>(
        &mut self,
        selector: impl Into<Selector<'selector>>,
        href: impl Into<String>,
    ) -> Result<&mut Self> {
        let reference = resolve(self.source, selector.into())?;
        let href = href.into();
        crate::link::validate_href(&href)?;
        self.links.insert(reference.get(), href);
        Ok(self)
    }

    /// Publishes every staged title/link effect as one fully reopened package.
    ///
    /// # Errors
    ///
    /// Returns an error for signed/encrypted input, non-compact package XML,
    /// invalid metadata, or semantic readback which differs from the request.
    pub fn commit(self) -> Result<Commit> {
        let title_changed = self.title_before != self.title_after;
        let link_changes = collect_link_changes(self.source, &self.links)?;
        if !title_changed && link_changes.is_empty() {
            return Ok(Commit::new(
                self.source,
                self.source.clone(),
                ChangeSet::default(),
            ));
        }

        let meta_xml = if title_changed {
            Some(stage_title(self.source, self.title_after.as_deref())?)
        } else {
            None
        };
        let mut content_xml = self.source.content_xml().to_owned();
        for change in link_changes.iter().rev() {
            let span = self
                .source
                .href_span(change.reference.get())
                .ok_or_else(|| invalid("ODM linked-section source span is missing"))?;
            content_xml = crate::link::replace_attribute_value(&content_xml, span, &change.after)?;
        }
        let snapshot = self.source.with_parts(&content_xml, meta_xml.as_deref())?;
        if snapshot.title() != self.title_after.as_deref() {
            return Err(invalid(
                "ODM transaction title readback differs from the request",
            ));
        }
        for change in &link_changes {
            let actual = snapshot
                .subdocuments()
                .get(change.reference.get())
                .ok_or_else(|| invalid("ODM transaction link disappeared during readback"))?;
            if actual.href() != change.after {
                return Err(invalid(
                    "ODM transaction link readback differs from the request",
                ));
            }
        }
        let title = title_changed.then_some(TitleChange {
            before: self.title_before,
            after: self.title_after,
        });
        Ok(Commit::new(
            self.source,
            snapshot,
            ChangeSet {
                title,
                links: link_changes,
            },
        ))
    }
}

/// The title effect of a unified transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TitleChange {
    before: Option<String>,
    after: Option<String>,
}

impl TitleChange {
    /// Returns the source title.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// Returns the published title.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }
}

/// One linked-section effect of a unified transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LinkChange {
    reference: Position,
    before: String,
    after: String,
}

impl LinkChange {
    /// Returns the checked reference position.
    #[must_use]
    pub const fn reference(&self) -> Position {
        self.reference
    }

    /// Returns the source target.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// Returns the published target.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// Ordered semantic effects retained by a unified patch.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ChangeSet {
    title: Option<TitleChange>,
    links: Vec<LinkChange>,
}

impl ChangeSet {
    /// Returns the title effect, when present.
    #[must_use]
    pub const fn title(&self) -> Option<&TitleChange> {
        self.title.as_ref()
    }

    /// Returns link effects in semantic reference order.
    #[must_use]
    pub fn links(&self) -> &[LinkChange] {
        &self.links
    }

    /// Returns whether this set contains no semantic effect.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.title.is_none() && self.links.is_empty()
    }
}

/// A fully validated unified transaction result.
pub struct Commit {
    snapshot: Master,
    patch: Patch,
}

impl Commit {
    fn new(source: &Master, snapshot: Master, changes: ChangeSet) -> Self {
        Self {
            patch: Patch {
                before: source.shared_bytes(),
                after: snapshot.shared_bytes(),
                changes,
            },
            snapshot,
        }
    }

    /// Returns the published immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Master {
        &self.snapshot
    }

    /// Returns the exact-source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit and returns its snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Master {
        self.snapshot
    }
}

/// An exact-source-checked reversible unified patch.
#[derive(Clone)]
pub struct Patch {
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
    changes: ChangeSet,
}

impl Patch {
    /// Returns the semantic effects.
    #[must_use]
    pub const fn changes(&self) -> &ChangeSet {
        &self.changes
    }

    /// Returns whether this patch applies to the exact source artifact.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Master) -> bool {
        source.as_bytes() == self.before.as_slice()
    }

    /// Applies this patch only to its exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` differs byte-for-byte.
    pub fn apply(&self, source: &Master) -> Result<Master> {
        if !self.is_applicable_to(source) {
            return Err(invalid("ODM unified patch source does not match"));
        }
        Master::from_shared_bytes(Arc::clone(&self.after))
    }

    /// Returns a patch restoring the exact source package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            changes: inverse_changes(&self.changes),
        }
    }

    /// Returns whether the exact package bytes are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_slice() == self.after.as_slice()
    }

    /// Merges effects derived from the same source when their writes agree.
    ///
    /// # Errors
    ///
    /// Returns typed conflicts for overlapping divergent writes, a source
    /// mismatch, or a validation failure while publishing the merged package.
    pub fn merge(&self, other: &Self) -> std::result::Result<Self, MergeError> {
        if self.before.as_slice() != other.before.as_slice() {
            return Err(MergeError::DifferentSource);
        }
        let conflicts =
            find_conflicts(&self.changes, &other.changes).map_err(MergeError::Invalid)?;
        if !conflicts.is_empty() {
            return Err(MergeError::Conflicts(ConflictSet::new(conflicts)));
        }
        let source =
            Master::from_shared_bytes(Arc::clone(&self.before)).map_err(MergeError::Invalid)?;
        let mut edit = source.edit();
        stage_changes(&mut edit, &self.changes).map_err(MergeError::Invalid)?;
        stage_changes(&mut edit, &other.changes).map_err(MergeError::Invalid)?;
        edit.commit()
            .map(|commit| commit.patch)
            .map_err(MergeError::Invalid)
    }

    /// Converts this patch to bounded canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error if either exact package exceeds the durable bounds or
    /// cannot be reopened without credentials.
    pub fn durable(&self) -> Result<DurablePatch> {
        DurablePatch::from_artifacts(self.before.as_slice(), self.after.as_slice())
    }
}

/// One semantic write target which prevented an automatic merge.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Conflict {
    /// Both patches write different titles.
    Title,
    /// Both patches write one link to different targets.
    Link(Position),
}

/// A unified-patch merge failure.
#[derive(Debug)]
#[non_exhaustive]
pub enum MergeError {
    /// The patches do not share exact source bytes.
    DifferentSource,
    /// Divergent writes overlap.
    Conflicts(ConflictSet<Conflict>),
    /// The merged candidate failed validation.
    Invalid(Error),
}

/// Explicit bounded undo/redo history for master snapshots.
pub struct History {
    inner: CoreHistory<Master>,
}

impl History {
    pub(crate) fn new(current: Master, limits: HistoryLimits) -> Self {
        Self {
            inner: CoreHistory::new(current, limits),
        }
    }

    /// Returns the current immutable snapshot.
    #[must_use]
    pub const fn current(&self) -> &Master {
        self.inner.current()
    }

    /// Records a commit only when history currently points at its source.
    ///
    /// # Errors
    ///
    /// Returns an error for stale lineage, size overflow, or exceeded bounds.
    pub fn record(&mut self, commit: &Commit) -> Result<Vec<Master>> {
        if !commit.patch.is_applicable_to(self.current()) {
            return Err(invalid(
                "ODM history commit source does not match current state",
            ));
        }
        let weight = u64::try_from(commit.patch.before.len())
            .ok()
            .and_then(|before| {
                u64::try_from(commit.patch.after.len())
                    .ok()
                    .and_then(|after| before.checked_add(after))
            })
            .ok_or_else(|| invalid("ODM history transition weight overflow"))?;
        self.inner
            .record(commit.snapshot.clone(), weight)
            .map_err(patch_wire_error)
    }

    /// Moves to the previous retained snapshot.
    pub fn undo(&mut self) -> bool {
        self.inner.undo()
    }

    /// Moves to the next retained snapshot.
    pub fn redo(&mut self) -> bool {
        self.inner.redo()
    }

    /// Returns whether undo is available.
    #[must_use]
    pub fn can_undo(&self) -> bool {
        self.inner.can_undo()
    }

    /// Returns whether redo is available.
    #[must_use]
    pub fn can_redo(&self) -> bool {
        self.inner.can_redo()
    }
}

/// Bounded reversible durable ODM patch.
#[derive(Clone)]
pub struct DurablePatch {
    inner: CorePatch<Reversible>,
}

impl DurablePatch {
    fn from_artifacts(before: &[u8], after: &[u8]) -> Result<Self> {
        Master::from_bytes(copy_bytes(before)?)?;
        Master::from_bytes(copy_bytes(after)?)?;
        let limits = durable_limits();
        let mut forward_blobs = BlobBundle::new(limits.blobs());
        let after_id = forward_blobs.insert(after).map_err(patch_wire_error)?;
        let mut reverse_blobs = BlobBundle::new(limits.blobs());
        let before_id = reverse_blobs.insert(before).map_err(patch_wire_error)?;
        let forward = durable_operation(limits, &before_id, &after_id)?;
        let inverse = durable_operation(limits, &after_id, &before_id)?;
        let inner = CorePatch::<Reversible>::new(
            limits,
            DURABLE_FORMAT,
            [ReversibleOperation::new(forward, inverse)],
            forward_blobs,
            reverse_blobs,
        )
        .map_err(patch_wire_error)?;
        Ok(Self { inner })
    }

    /// Parses canonical deterministic JSON and validates both ODM artifacts.
    ///
    /// # Errors
    ///
    /// Returns an error for non-canonical, foreign, excessive, or invalid data.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        let inner = CorePatch::<Reversible>::from_deterministic_json(bytes, durable_limits())
            .map_err(patch_wire_error)?;
        validate_reversible(&inner)?;
        Ok(Self { inner })
    }

    /// Serializes canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error when bounded serialization fails.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        self.inner.to_deterministic_json().map_err(patch_wire_error)
    }

    /// Applies this durable patch to its exact source artifact.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source bytes or invalid target bytes.
    pub fn apply(&self, source: &Master) -> Result<Master> {
        let inverse = self.inner.inverse();
        if source.as_bytes() != durable_direction(&inverse)?.target_bytes {
            return Err(invalid("ODM durable patch source does not match"));
        }
        master_from_target(&self.inner)
    }

    /// Returns the exact durable inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            inner: self.inner.inverse(),
        }
    }

    /// Permanently removes inverse material.
    #[must_use]
    pub fn seal(self) -> SealedPatch {
        SealedPatch {
            inner: self.inner.seal(),
        }
    }
}

/// Forward-only durable ODM patch.
#[derive(Clone)]
pub struct SealedPatch {
    inner: CorePatch<ForwardOnly>,
}

impl SealedPatch {
    /// Parses a canonical forward-only durable patch.
    ///
    /// # Errors
    ///
    /// Returns an error for non-canonical, foreign, excessive, or invalid data.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        let inner = CorePatch::<ForwardOnly>::from_deterministic_json(bytes, durable_limits())
            .map_err(patch_wire_error)?;
        validate_sealed(&inner)?;
        Ok(Self { inner })
    }

    /// Serializes canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error when bounded serialization fails.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        self.inner.to_deterministic_json().map_err(patch_wire_error)
    }

    /// Applies after checking the retained SHA-256 source precondition.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source bytes or invalid target bytes.
    pub fn apply(&self, source: &Master) -> Result<Master> {
        let direction = durable_direction(&self.inner)?;
        if BlobId::of(source.as_bytes()).as_hex() != direction.source_id {
            return Err(invalid("ODM durable patch source does not match"));
        }
        Master::from_bytes(copy_bytes(direction.target_bytes)?)
    }
}

fn resolve(source: &Master, selector: Selector<'_>) -> Result<Position> {
    match selector {
        Selector::Position(position) => source
            .subdocuments()
            .get(position.get())
            .map(|_| position)
            .ok_or_else(|| invalid("ODM subdocument selector is out of bounds")),
        Selector::Section(name) => source
            .subdocuments()
            .iter()
            .position(|reference| reference.section() == name.as_ref())
            .map(Position::new)
            .ok_or_else(|| invalid("ODM linked section name was not found")),
    }
}

fn collect_link_changes(
    source: &Master,
    staged: &BTreeMap<usize, String>,
) -> Result<Vec<LinkChange>> {
    let mut changes = Vec::new();
    changes
        .try_reserve(staged.len())
        .map_err(|allocation| Error::Allocation {
            resource: "ODM unified link changes",
            source: allocation,
        })?;
    for (&index, after) in staged {
        let before = source
            .subdocuments()
            .get(index)
            .ok_or_else(|| invalid("ODM staged reference disappeared"))?
            .href();
        if before != after {
            changes.push(LinkChange {
                reference: Position::new(index),
                before: before.to_owned(),
                after: after.clone(),
            });
        }
    }
    Ok(changes)
}

fn stage_title(source: &Master, after: Option<&str>) -> Result<String> {
    let source_xml = source
        .meta_xml()?
        .ok_or_else(|| invalid("ODM title editing requires an existing UTF-8 meta.xml part"))?;
    let metadata = OdfMetadata::from_xml(&source_xml)?;
    let mut changed = litchi_core::Metadata::from(metadata.clone());
    changed.title = after.map(str::to_owned);
    let patch = MetaXmlPatch::preserve_all().diff_simple_fields(&metadata, &changed);
    patch_meta_xml(&source_xml, &patch)?
        .ok_or_else(|| invalid("ODM title editing requires an office:meta container"))
}

fn inverse_changes(changes: &ChangeSet) -> ChangeSet {
    ChangeSet {
        title: changes.title.as_ref().map(|change| TitleChange {
            before: change.after.clone(),
            after: change.before.clone(),
        }),
        links: changes
            .links
            .iter()
            .map(|change| LinkChange {
                reference: change.reference,
                before: change.after.clone(),
                after: change.before.clone(),
            })
            .collect(),
    }
}

fn find_conflicts(left: &ChangeSet, right: &ChangeSet) -> Result<Vec<Conflict>> {
    let mut conflicts = Vec::new();
    conflicts
        .try_reserve(left.links.len().min(right.links.len()).saturating_add(1))
        .map_err(|allocation| Error::Allocation {
            resource: "ODM merge conflicts",
            source: allocation,
        })?;
    if let (Some(left), Some(right)) = (&left.title, &right.title)
        && left.after != right.after
    {
        conflicts.push(Conflict::Title);
    }
    let mut right_links = HashMap::new();
    right_links
        .try_reserve(right.links.len())
        .map_err(|allocation| Error::Allocation {
            resource: "ODM merge link index",
            source: allocation,
        })?;
    for change in &right.links {
        right_links.insert(change.reference.get(), change.after.as_str());
    }
    for change in &left.links {
        if right_links
            .get(&change.reference.get())
            .is_some_and(|after| *after != change.after)
        {
            conflicts.push(Conflict::Link(change.reference));
        }
    }
    Ok(conflicts)
}

fn stage_changes(edit: &mut Edit<'_>, changes: &ChangeSet) -> Result<()> {
    if let Some(title) = &changes.title {
        if let Some(after) = &title.after {
            edit.set_title(after.clone())?;
        } else {
            edit.clear_title();
        }
    }
    for link in &changes.links {
        edit.set_link(link.reference, link.after.clone())?;
    }
    Ok(())
}

struct DurableDirection<'a> {
    source_id: &'a str,
    target_id: &'a str,
    target_bytes: &'a [u8],
}

fn durable_limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(1, MAX_PACKAGE_BYTES, MAX_PACKAGE_BYTES),
        MAX_WIRE_JSON_BYTES,
        1,
        4,
        4_096,
        16_384,
    )
}

fn durable_operation(
    limits: PatchLimits,
    source: &BlobId,
    target: &BlobId,
) -> Result<PatchOperation> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        SOURCE_PRECONDITION.to_string(),
        Value::String(source.as_hex()),
    );
    PatchOperation::new(
        limits,
        DURABLE_OPERATION,
        DURABLE_TARGET,
        preconditions,
        Value::String(target.as_hex()),
    )
    .map_err(patch_wire_error)
}

fn durable_direction<Mode>(patch: &CorePatch<Mode>) -> Result<DurableDirection<'_>> {
    if patch.format() != DURABLE_FORMAT || patch.operations().len() != 1 {
        return Err(invalid("invalid ODM durable patch vocabulary"));
    }
    let operation = &patch.operations()[0];
    if operation.op != DURABLE_OPERATION
        || operation.target != DURABLE_TARGET
        || operation.preconditions.len() != 1
        || patch.blobs().len() != 1
    {
        return Err(invalid("invalid ODM durable patch vocabulary"));
    }
    let source_id = operation
        .preconditions
        .get(SOURCE_PRECONDITION)
        .and_then(Value::as_str)
        .filter(|value| is_digest(value))
        .ok_or_else(|| invalid("invalid ODM durable patch vocabulary"))?;
    let target_id = operation
        .value
        .as_str()
        .filter(|value| is_digest(value))
        .ok_or_else(|| invalid("invalid ODM durable patch vocabulary"))?;
    let blob_id = patch
        .blobs()
        .ids()
        .next()
        .filter(|id| id.as_hex() == target_id)
        .ok_or_else(|| invalid("invalid ODM durable patch vocabulary"))?;
    let target_bytes = patch
        .blobs()
        .get(blob_id)
        .ok_or_else(|| invalid("invalid ODM durable patch vocabulary"))?;
    Ok(DurableDirection {
        source_id,
        target_id,
        target_bytes,
    })
}

fn validate_reversible(patch: &CorePatch<Reversible>) -> Result<()> {
    let forward = durable_direction(patch)?;
    let inverse = patch.inverse();
    let reverse = durable_direction(&inverse)?;
    if forward.source_id != reverse.target_id || forward.target_id != reverse.source_id {
        return Err(invalid("invalid ODM durable patch vocabulary"));
    }
    Master::from_bytes(copy_bytes(forward.target_bytes)?)?;
    Master::from_bytes(copy_bytes(reverse.target_bytes)?)?;
    Ok(())
}

fn validate_sealed(patch: &CorePatch<ForwardOnly>) -> Result<()> {
    Master::from_bytes(copy_bytes(durable_direction(patch)?.target_bytes)?)?;
    Ok(())
}

fn master_from_target<Mode>(patch: &CorePatch<Mode>) -> Result<Master> {
    Master::from_bytes(copy_bytes(durable_direction(patch)?.target_bytes)?)
}

fn is_digest(value: &str) -> bool {
    value.len() == 64
        && value
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
}

fn copy_bytes(source: &[u8]) -> Result<Vec<u8>> {
    if source.len() > MAX_PACKAGE_BYTES {
        return Err(invalid("ODM durable package exceeds the 256 MiB limit"));
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(source.len())
        .map_err(|allocation| Error::Allocation {
            resource: "ODM durable package",
            source: allocation,
        })?;
    bytes.extend_from_slice(source);
    Ok(bytes)
}

fn patch_wire_error(source: litchi_core::PatchError) -> Error {
    let message = format!("invalid ODM durable patch: {source}");
    drop(source);
    Error::InvalidFormat(message)
}

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(message.to_owned())
}
