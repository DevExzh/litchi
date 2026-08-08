//! Reachable-story inventory, checksum verification, and atomic publication.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_ooxml_common::custom_xml;
use litchi_opc::{OpcPackage, PackURI};

use crate::package::story::{self, StoryKind, StoryLimits, StoryTopology};
use crate::{Error, Package, Result};

use super::patch::Gate;
use super::{BindingFlavor, Checksum, ChecksumStatus, Commit, Limits, Snapshot, Transaction};

const SIGNATURE_TOKEN_MAGIC: &[u8] = b"litchi.docx.sdt.signatures.v1\0";

/// Aggregate bounds for package-wide content-control work.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PackageLimits {
    /// Per-story semantic and exact-source limits.
    pub controls: Limits,
    /// Reachable Word story-graph limits.
    pub stories: StoryLimits,
    /// Maximum aggregate active content controls.
    pub max_content_controls: usize,
    /// Maximum aggregate active data bindings across all reachable stories.
    pub max_bindings: usize,
    /// Maximum entries materialized by one checksum verification report.
    pub max_report_entries: usize,
    /// Maximum unique Custom XML payloads processed for CRCs.
    pub max_crc_parts: usize,
    /// Maximum aggregate bytes processed once across unique CRC payloads.
    pub max_crc_bytes: usize,
    /// Maximum signature graph bytes retained for stale checks.
    pub max_signature_bytes: usize,
    /// Maximum canonical Custom XML graph metadata retained for stale checks.
    pub max_custom_graph_bytes: usize,
    /// Maximum Custom XML relationship occurrences inspected before discovery.
    pub max_custom_relationships: usize,
    /// Maximum story mutations in one transaction.
    pub max_mutations: usize,
    /// Maximum aggregate rewritten story bytes.
    pub max_output_bytes: usize,
}

impl Default for PackageLimits {
    fn default() -> Self {
        Self {
            controls: Limits::default(),
            stories: StoryLimits::default(),
            max_content_controls: 65_536,
            max_bindings: 65_536,
            max_report_entries: 65_536,
            max_crc_parts: 4_096,
            max_crc_bytes: 256 * 1024 * 1024,
            max_signature_bytes: 16 * 1024 * 1024,
            max_custom_graph_bytes: 64 * 1024 * 1024,
            max_custom_relationships: 65_536,
            max_mutations: 65_536,
            max_output_bytes: 256 * 1024 * 1024,
        }
    }
}

impl PackageLimits {
    fn validate(&self) -> Result<()> {
        self.controls.validate()?;
        if [
            self.max_content_controls,
            self.max_bindings,
            self.max_report_entries,
            self.max_crc_parts,
            self.max_crc_bytes,
            self.max_signature_bytes,
            self.max_custom_graph_bytes,
            self.max_custom_relationships,
            self.max_mutations,
            self.max_output_bytes,
        ]
        .contains(&0)
        {
            return Err(Error::Invalid(
                "package content-control limits must be nonzero".into(),
            ));
        }
        Ok(())
    }
}

/// One reachable WordprocessingML story with an exact content-control snapshot.
#[derive(Debug, Clone)]
pub struct Story {
    part: PackURI,
    content_type: String,
    kind: StoryKind,
    snapshot: Snapshot,
}

impl Story {
    /// Canonical owning part.
    #[must_use]
    pub const fn part(&self) -> &PackURI {
        &self.part
    }

    /// Declared part content type.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Semantic story role.
    #[must_use]
    pub const fn kind(&self) -> StoryKind {
        self.kind
    }

    /// Exact story snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }
}

#[derive(Debug, Clone)]
struct Store {
    id: String,
    part: PackURI,
    content_type: String,
    data: Arc<Vec<u8>>,
    props_part: Option<PackURI>,
}

/// Immutable package-bound inventory and stale-source precondition.
#[derive(Debug, Clone)]
pub struct PackageSnapshot {
    stories: Arc<[Story]>,
    topology: StoryTopology,
    stores: Arc<[Store]>,
    store_index: Arc<HashMap<String, Vec<usize>>>,
    custom_graph: Arc<[u8]>,
    signatures: Arc<[u8]>,
    limits: PackageLimits,
}

impl PackageSnapshot {
    fn capture(package: &OpcPackage, limits: PackageLimits) -> Result<Self> {
        limits.validate()?;
        let inventory = story::capture(package, limits.stories.clone())?;
        let mut stories = Vec::new();
        stories
            .try_reserve_exact(inventory.stories().len())
            .map_err(alloc("content-control package stories"))?;
        let mut controls = 0usize;
        let mut bindings = 0usize;
        for located in inventory.stories() {
            let snapshot = Snapshot::from_package(located.source_arc(), limits.controls.clone())?;
            controls = controls
                .checked_add(snapshot.occurrences().len())
                .ok_or_else(|| invalid("package content-control count overflow"))?;
            if controls > limits.max_content_controls {
                return Err(invalid(
                    "package content-control count exceeds configured limit",
                ));
            }
            for occurrence in snapshot.inventory().occurrences() {
                bindings = bindings
                    .checked_add(occurrence.control().data_bindings().len())
                    .ok_or_else(|| invalid("package data-binding count overflow"))?;
                if bindings > limits.max_bindings {
                    return Err(invalid(
                        "package data-binding count exceeds configured limit",
                    ));
                }
            }
            stories.push(Story {
                part: located.part().clone(),
                content_type: located.content_type().to_owned(),
                kind: located.kind(),
                snapshot,
            });
        }
        let (stores, custom_graph) = capture_stores(package, &limits)?;
        let store_index = index_stores(&stores)?;
        let signatures = signature_token(package, limits.max_signature_bytes)?;
        Ok(Self {
            stories: Arc::from(stories.into_boxed_slice()),
            topology: inventory.topology().clone(),
            stores,
            store_index,
            custom_graph,
            signatures,
            limits,
        })
    }

    /// Reachable stories, main document first and then canonical part order.
    #[must_use]
    pub fn stories(&self) -> &[Story] {
        &self.stories
    }

    /// Opaque canonical topology token captured with these stories.
    #[must_use]
    pub const fn topology(&self) -> &StoryTopology {
        &self.topology
    }

    /// Start a bounded package-wide transaction.
    pub fn edit(&self) -> Result<PackageTransaction> {
        let mut transactions = Vec::new();
        transactions
            .try_reserve_exact(self.stories.len())
            .map_err(alloc("content-control story transactions"))?;
        for story in self.stories.iter() {
            transactions.push(story.snapshot.edit());
        }
        Ok(PackageTransaction {
            base: self.clone(),
            transactions,
            mutations: HashSet::new(),
            crc_memo: HashMap::new(),
            crc_bytes: 0,
        })
    }

    /// Verify every declared checksum without executing XPath or interpreting XML.
    ///
    /// Snapshot capture rejects duplicate Custom XML datastore item GUIDs, so
    /// verification never selects an arbitrary payload for an ambiguous ID.
    pub fn verify_checksums(&self) -> Result<Vec<ChecksumEntry>> {
        verify(self, None)
    }

    fn story_index(&self, part: &PackURI) -> Result<usize> {
        self.stories
            .iter()
            .position(|story| story.part == *part)
            .ok_or_else(|| Error::PartNotFound(part.as_str().to_owned()))
    }

    fn store_indices(&self, id: &str) -> &[usize] {
        self.store_index
            .get(&id.to_ascii_lowercase())
            .map_or(&[], Vec::as_slice)
    }

    fn unique_store(&self, id: &str) -> Result<&Store> {
        let indices = self.store_indices(id);
        if indices.is_empty() {
            return Err(Error::PartNotFound(format!("Custom XML itemID '{id}'")));
        }
        if indices.len() != 1 {
            return Err(invalid(format!("Custom XML itemID '{id}' is ambiguous")));
        }
        self.stores
            .get(indices[0])
            .ok_or_else(|| invalid("Custom XML GUID index is corrupt"))
    }
}

/// Package-aware checksum outcome for one active binding.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PackageChecksumStatus {
    /// No checksum was declared.
    Absent,
    /// Checksum lexical text is malformed but preserved.
    Malformed(Box<str>),
    /// No reachable Custom XML store owns the declared GUID.
    MissingStore,
    /// The exact payload bytes match.
    Matches,
    /// The declared and computed values differ.
    Mismatch {
        /// Declared value.
        expected: Checksum,
        /// Value over current exact payload bytes.
        actual: Checksum,
    },
}

/// One source-bound package checksum report entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChecksumEntry {
    part: PackURI,
    occurrence: usize,
    binding: usize,
    flavor: BindingFlavor,
    control_id: Option<u32>,
    store_item_id: String,
    store_part: Option<PackURI>,
    status: PackageChecksumStatus,
}

impl ChecksumEntry {
    /// Owning story part.
    #[must_use]
    pub const fn part(&self) -> &PackURI {
        &self.part
    }

    /// Source-order occurrence within the story.
    #[must_use]
    pub const fn occurrence(&self) -> usize {
        self.occurrence
    }

    /// Source-order binding index within the content control.
    #[must_use]
    pub const fn binding(&self) -> usize {
        self.binding
    }

    /// Exact owner vocabulary of the reported binding.
    #[must_use]
    pub const fn flavor(&self) -> BindingFlavor {
        self.flavor
    }

    /// Optional producer ID, which is not used as transaction identity.
    #[must_use]
    pub const fn control_id(&self) -> Option<u32> {
        self.control_id
    }

    /// Declared Custom XML GUID.
    #[must_use]
    pub fn store_item_id(&self) -> &str {
        &self.store_item_id
    }

    /// Resolved data part when unique.
    #[must_use]
    pub const fn store_part(&self) -> Option<&PackURI> {
        self.store_part.as_ref()
    }

    /// Verification outcome. CRC equality is advisory, not authentication.
    #[must_use]
    pub const fn status(&self) -> &PackageChecksumStatus {
        &self.status
    }
}

/// Package-wide edit queue keyed by story part and source ordinal.
#[derive(Debug, Clone)]
pub struct PackageTransaction {
    base: PackageSnapshot,
    transactions: Vec<Transaction>,
    mutations: HashSet<(usize, usize, MutationKind)>,
    crc_memo: HashMap<PackURI, Checksum>,
    crc_bytes: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum MutationKind {
    Checksum(usize),
    Formatting,
}

impl PackageTransaction {
    /// Exact package snapshot captured for stale checks.
    #[must_use]
    pub const fn source(&self) -> &PackageSnapshot {
        &self.base
    }

    /// Set, repair, or remove one checksum by stable package occurrence.
    pub fn set_checksum(
        &mut self,
        part: &PackURI,
        occurrence: usize,
        value: Option<Checksum>,
    ) -> Result<&mut Self> {
        let story = self.base.story_index(part)?;
        let semantic = self.base.stories[story]
            .snapshot
            .inventory()
            .occurrences()
            .get(occurrence)
            .ok_or(Error::OutOfBounds {
                object: "content-control occurrence",
                index: occurrence,
                len: self.base.stories[story]
                    .snapshot
                    .inventory()
                    .occurrences()
                    .len(),
            })?;
        let bindings = semantic.control().data_bindings();
        let binding = bindings
            .iter()
            .position(|binding| binding.flavor() == BindingFlavor::Core)
            .or_else(|| (!bindings.is_empty()).then_some(0))
            .ok_or_else(|| invalid("content control has no active binding"))?;
        self.set_binding_checksum(part, occurrence, binding, value)
    }

    /// Set, repair, or remove one exact source-order binding checksum.
    pub fn set_binding_checksum(
        &mut self,
        part: &PackURI,
        occurrence: usize,
        binding: usize,
        value: Option<Checksum>,
    ) -> Result<&mut Self> {
        let story = self.base.story_index(part)?;
        let kind = MutationKind::Checksum(binding);
        let reserved = self.preflight_mutation(story, occurrence, kind)?;
        self.transactions[story].set_binding_checksum(occurrence, binding, value)?;
        self.finish_mutation(story, occurrence, kind, reserved);
        Ok(self)
    }

    /// Refresh one checksum from its exact current Custom XML part bytes.
    pub fn refresh_checksum(&mut self, part: &PackURI, occurrence: usize) -> Result<&mut Self> {
        let story = self.base.story_index(part)?;
        let semantic = self.base.stories[story]
            .snapshot
            .inventory()
            .occurrences()
            .get(occurrence)
            .ok_or(Error::OutOfBounds {
                object: "content-control occurrence",
                index: occurrence,
                len: self.base.stories[story]
                    .snapshot
                    .inventory()
                    .occurrences()
                    .len(),
            })?;
        let binding = semantic
            .control()
            .data_binding()
            .ok_or_else(|| invalid("content control has no data binding"))?;
        let store = self.base.unique_store(binding.store_item_id())?;
        let store_part = store.part.clone();
        let store_data = store.data.clone();
        let checksum = self.checksum_for_store(&store_part, store_data.as_slice())?;
        self.set_checksum(part, occurrence, Some(checksum))
    }

    /// Refresh every present checksum, preserving checksum absence.
    ///
    /// Each unique Custom XML part is processed once, irrespective of the
    /// number of bindings that reference it.
    pub fn refresh_checksums(&mut self) -> Result<&mut Self> {
        let mut target_count = 0usize;
        let mut additional_mutations = 0usize;
        let mut story_additions = Vec::new();
        story_additions
            .try_reserve_exact(self.transactions.len())
            .map_err(alloc("content-control story mutation reservations"))?;
        story_additions.resize(self.transactions.len(), 0usize);
        for story in 0..self.base.stories.len() {
            for semantic in self.base.stories[story].snapshot.inventory().occurrences() {
                for (binding, value) in semantic.control().data_bindings().iter().enumerate() {
                    if value.checksum_value().is_none() {
                        continue;
                    }
                    target_count = target_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("checksum refresh target count overflow"))?;
                    self.transactions[story]
                        .validate_checksum_target(semantic.ordinal(), binding)?;
                    if !self.mutations.contains(&(
                        story,
                        semantic.ordinal(),
                        MutationKind::Checksum(binding),
                    )) {
                        additional_mutations =
                            additional_mutations.checked_add(1).ok_or_else(|| {
                                invalid("content-control package mutation count overflow")
                            })?;
                        story_additions[story] = story_additions[story]
                            .checked_add(1)
                            .ok_or_else(|| invalid("story mutation reservation overflow"))?;
                    }
                }
            }
        }
        let total_mutations = self
            .mutations
            .len()
            .checked_add(additional_mutations)
            .ok_or_else(|| invalid("content-control package mutation count overflow"))?;
        if total_mutations > self.base.limits.max_mutations {
            return Err(invalid(
                "content-control package mutation count exceeds configured limit",
            ));
        }
        self.mutations
            .try_reserve(additional_mutations)
            .map_err(alloc("content-control package mutations"))?;
        for (transaction, additional) in self
            .transactions
            .iter_mut()
            .zip(story_additions.into_iter())
        {
            transaction.try_reserve_edits(additional)?;
        }

        let mut plans = Vec::<(usize, usize, usize, Checksum)>::new();
        plans
            .try_reserve_exact(target_count)
            .map_err(alloc("content-control checksum refresh plans"))?;
        for story_index in 0..self.base.stories.len() {
            let occurrence_count = self.base.stories[story_index]
                .snapshot
                .inventory()
                .occurrences()
                .len();
            for occurrence_index in 0..occurrence_count {
                let candidates = {
                    let semantic = &self.base.stories[story_index]
                        .snapshot
                        .inventory()
                        .occurrences()[occurrence_index];
                    let bindings = semantic.control().data_bindings();
                    let mut candidates = Vec::new();
                    candidates
                        .try_reserve_exact(bindings.len())
                        .map_err(alloc("content-control checksum refresh candidates"))?;
                    for (binding_index, binding) in bindings.iter().enumerate() {
                        if binding.checksum_value().is_none() {
                            continue;
                        }
                        let store = self.base.unique_store(binding.store_item_id())?;
                        candidates.push((
                            semantic.ordinal(),
                            binding_index,
                            store.part.clone(),
                            store.data.clone(),
                        ));
                    }
                    candidates
                };
                for (ordinal, binding, store_part, store_data) in candidates {
                    let checksum = self.checksum_for_store(&store_part, store_data.as_slice())?;
                    plans
                        .try_reserve(1)
                        .map_err(alloc("content-control checksum refresh plans"))?;
                    plans.push((story_index, ordinal, binding, checksum));
                }
            }
        }
        for (story, occurrence, binding, checksum) in plans {
            let kind = MutationKind::Checksum(binding);
            let reserved = self.preflight_mutation(story, occurrence, kind)?;
            self.transactions[story].set_binding_checksum(occurrence, binding, Some(checksum))?;
            self.finish_mutation(story, occurrence, kind, reserved);
        }
        Ok(self)
    }

    fn checksum_for_store(&mut self, part: &PackURI, data: &[u8]) -> Result<Checksum> {
        if let Some(value) = self.crc_memo.get(part) {
            return Ok(value.clone());
        }
        if self.crc_memo.len() >= self.base.limits.max_crc_parts {
            return Err(invalid("unique CRC part limit exceeded"));
        }
        let crc_bytes = self
            .crc_bytes
            .checked_add(data.len())
            .ok_or_else(|| invalid("aggregate CRC byte count overflow"))?;
        if crc_bytes > self.base.limits.max_crc_bytes {
            return Err(invalid("aggregate CRC byte limit exceeded"));
        }
        self.crc_memo
            .try_reserve(1)
            .map_err(alloc("content-control CRC memo"))?;
        let value = Checksum::compute(data, &self.base.limits.controls)?;
        self.crc_memo.insert(part.clone(), value.clone());
        self.crc_bytes = crc_bytes;
        Ok(value)
    }

    /// Set or remove the formatting exception on one exact lock.
    pub fn set_formatting_allowed(
        &mut self,
        part: &PackURI,
        occurrence: usize,
        value: Option<super::FormattingAllowed>,
    ) -> Result<&mut Self> {
        let story = self.base.story_index(part)?;
        let reserved = self.preflight_mutation(story, occurrence, MutationKind::Formatting)?;
        self.transactions[story].set_formatting_allowed(occurrence, value)?;
        self.finish_mutation(story, occurrence, MutationKind::Formatting, reserved);
        Ok(self)
    }

    /// Materialize and fully reparse every changed story.
    pub fn commit(&self) -> Result<PackageCommit> {
        let mut commits = Vec::new();
        commits
            .try_reserve_exact(self.transactions.len())
            .map_err(alloc("content-control story commits"))?;
        let mut output = 0usize;
        for transaction in &self.transactions {
            let commit = transaction.commit()?;
            if commit.changed() {
                output = output
                    .checked_add(commit.snapshot().source().len())
                    .ok_or_else(|| invalid("aggregate content-control output overflow"))?;
                if output > self.base.limits.max_output_bytes {
                    return Err(invalid("aggregate content-control output limit exceeded"));
                }
            }
            commits.push(commit);
        }
        let mut after = Vec::new();
        after
            .try_reserve_exact(commits.len())
            .map_err(alloc("content-control package targets"))?;
        after.extend(commits.iter().map(|commit| commit.snapshot().clone()));
        let changed = commits.iter().any(Commit::changed);
        let after_signatures = if changed {
            unsigned_signature_token()?
        } else {
            self.base.signatures.clone()
        };
        let patch = PackagePatch {
            before: self.base.clone(),
            after: Arc::from(after.into_boxed_slice()),
            after_signatures,
            gate: Gate::new(),
        };
        Ok(PackageCommit {
            commits: Arc::from(commits.into_boxed_slice()),
            patch,
        })
    }

    fn preflight_mutation(
        &mut self,
        story: usize,
        occurrence: usize,
        kind: MutationKind,
    ) -> Result<bool> {
        if self.mutations.contains(&(story, occurrence, kind)) {
            return Ok(false);
        }
        if self.mutations.len() >= self.base.limits.max_mutations {
            return Err(invalid("content-control mutation limit exceeded"));
        }
        self.mutations
            .try_reserve(1)
            .map_err(alloc("content-control mutation index"))?;
        Ok(true)
    }

    fn finish_mutation(
        &mut self,
        story: usize,
        occurrence: usize,
        kind: MutationKind,
        reserved: bool,
    ) {
        if reserved {
            self.mutations.insert((story, occurrence, kind));
        }
    }
}

/// Prepared package publication.
#[derive(Debug, Clone)]
pub struct PackageCommit {
    commits: Arc<[Commit]>,
    patch: PackagePatch,
}

impl PackageCommit {
    /// Whether one or more story sources change.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.commits.iter().any(Commit::changed)
    }

    /// Fully reparsed per-story commits in package story order.
    #[must_use]
    pub fn stories(&self) -> &[Commit] {
        &self.commits
    }

    /// Package-bound retry-safe patch.
    #[must_use]
    pub const fn patch(&self) -> &PackagePatch {
        &self.patch
    }
}

/// Retry-safe, topology- and source-checked package publication.
#[derive(Debug, Clone)]
pub struct PackagePatch {
    before: PackageSnapshot,
    after: Arc<[Snapshot]>,
    after_signatures: Arc<[u8]>,
    gate: Gate,
}

impl PackagePatch {
    /// Whether every story byte remains identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before
            .stories
            .iter()
            .zip(self.after.iter())
            .all(|(before, after)| before.snapshot.source() == after.source())
    }

    /// Whether this patch or a clone was published successfully.
    #[must_use]
    pub fn is_applied(&self) -> bool {
        self.gate.is_applied()
    }

    /// Construct a fresh inverse publication for the exact result signature state.
    ///
    /// The inverse retains the same topology and Custom XML preconditions but
    /// requires the exact candidate story bytes. Signatures are never restored
    /// by an inverse story edit; a re-signed package must be explicitly
    /// unsigned again before a changed inverse can be published.
    pub fn inverse(&self) -> Result<Self> {
        let mut stories = Vec::new();
        stories
            .try_reserve_exact(self.before.stories.len())
            .map_err(alloc("inverse content-control stories"))?;
        for (story, snapshot) in self.before.stories.iter().zip(self.after.iter()) {
            stories.push(Story {
                part: story.part.clone(),
                content_type: story.content_type.clone(),
                kind: story.kind,
                snapshot: snapshot.clone(),
            });
        }
        let target = PackageSnapshot {
            stories: Arc::from(stories.into_boxed_slice()),
            topology: self.before.topology.clone(),
            stores: self.before.stores.clone(),
            store_index: self.before.store_index.clone(),
            custom_graph: self.before.custom_graph.clone(),
            signatures: self.after_signatures.clone(),
            limits: self.before.limits.clone(),
        };
        let mut original = Vec::new();
        original
            .try_reserve_exact(self.before.stories.len())
            .map_err(alloc("inverse content-control targets"))?;
        original.extend(
            self.before
                .stories
                .iter()
                .map(|story| story.snapshot.clone()),
        );
        let result_signatures = if self.is_noop() {
            self.before.signatures.clone()
        } else {
            self.after_signatures.clone()
        };
        Ok(Self {
            before: target,
            after: Arc::from(original.into_boxed_slice()),
            after_signatures: result_signatures,
            gate: Gate::new(),
        })
    }
}

impl Package {
    /// Capture all reachable content-control stories and Custom XML preconditions.
    pub fn content_control_snapshot(&self) -> Result<PackageSnapshot> {
        self.ensure_story_opc_current("content_control_snapshot")?;
        PackageSnapshot::capture(self.opc_package(), PackageLimits::default())
    }

    /// Capture with explicit aggregate resource limits.
    pub fn content_control_snapshot_with_limits(
        &self,
        limits: PackageLimits,
    ) -> Result<PackageSnapshot> {
        self.ensure_story_opc_current("content_control_snapshot_with_limits")?;
        PackageSnapshot::capture(self.opc_package(), limits)
    }

    /// Verify declared checksums across every reachable story.
    pub fn verify_content_control_checksums(&self) -> Result<Vec<ChecksumEntry>> {
        self.ensure_story_opc_current("verify_content_control_checksums")?;
        PackageSnapshot::capture(self.opc_package(), PackageLimits::default())?.verify_checksums()
    }

    /// Publish a prepared package transaction.
    pub fn apply_content_controls(&mut self, commit: &PackageCommit) -> Result<()> {
        self.ensure_story_opc_current("apply_content_controls")?;
        self.apply_content_control_patch(commit.patch())
    }

    /// Publish one exact package patch atomically.
    pub fn apply_content_control_patch(&mut self, patch: &PackagePatch) -> Result<()> {
        // Guard before claiming the shared retry-safe gate. A pending facade
        // edit must not consume even an exact no-op patch.
        self.ensure_story_opc_current("apply_content_control_patch")?;
        let current = PackageSnapshot::capture(self.opc_package(), patch.before.limits.clone())?;
        if patch.is_noop() {
            require_current(&patch.before, &current, true)?;
            let claim = patch.gate.claim("content-control package patch")?;
            claim.finalize();
            return Ok(());
        }
        require_current(&patch.before, &current, false)?;
        if !is_unsigned_signature_token(&current.signatures) {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "apply_content_control_patch",
                reason: "changed content controls cannot preserve package signatures; call Package::unsign explicitly first",
            });
        }
        let claim = patch.gate.claim("content-control package patch")?;

        let topology = patch.before.topology.clone();
        let mut before_stories = Vec::new();
        before_stories
            .try_reserve_exact(patch.before.stories.len())
            .map_err(alloc("content-control story before-images"))?;
        for story in patch.before.stories.iter() {
            let source = match story.snapshot.source_owner() {
                super::snapshot::Source::Package(value) => value,
                super::snapshot::Source::Detached(_) => {
                    return Err(invalid("package patch retained a detached story source"));
                },
            };
            before_stories.push((story.part.clone(), source));
        }
        let after = patch.after.clone();
        let limits = patch.before.limits.clone();
        self.edit_semantic_opc("apply_content_control_patch", move |candidate| {
            let staged = PackageSnapshot::capture(candidate, limits.clone())?;
            if staged.topology != topology {
                return Err(invalid("content-control story topology is stale"));
            }
            if staged.custom_graph != current.custom_graph
                || !same_stores(&staged.stores, &current.stores)
            {
                return Err(invalid("Custom XML checksum source is stale"));
            }
            for ((part, before), replacement) in before_stories.iter().zip(after.iter()) {
                if candidate.get_part(part)?.blob() != before.as_slice() {
                    return Err(invalid("content-control story source is stale"));
                }
                if replacement.source() != before.as_slice() {
                    candidate.get_part_mut(part)?.set_blob_shared(
                        match replacement.source_owner() {
                            super::snapshot::Source::Package(value) => value,
                            super::snapshot::Source::Detached(value) => {
                                Arc::new(value.as_ref().to_vec())
                            },
                        },
                    );
                }
            }
            let reparsed = PackageSnapshot::capture(candidate, limits)?;
            verify(&reparsed, Some(true))?;
            Ok(())
        })?;
        claim.finalize();
        Ok(())
    }
}

fn verify(snapshot: &PackageSnapshot, require_matches: Option<bool>) -> Result<Vec<ChecksumEntry>> {
    let report_entries = snapshot.stories.iter().try_fold(0usize, |total, story| {
        story
            .snapshot
            .inventory()
            .occurrences()
            .iter()
            .try_fold(total, |total, occurrence| {
                total
                    .checked_add(occurrence.control().data_bindings().len())
                    .ok_or_else(|| invalid("checksum report entry count overflow"))
            })
    })?;
    if report_entries > snapshot.limits.max_report_entries {
        return Err(invalid(
            "checksum report entry count exceeds configured limit",
        ));
    }
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(report_entries)
        .map_err(alloc("checksum report"))?;
    let mut memo = HashMap::<PackURI, Checksum>::new();
    let mut total = 0usize;
    for story in snapshot.stories.iter() {
        for occurrence in story.snapshot.inventory().occurrences() {
            for (binding_index, binding) in occurrence.control().data_bindings().iter().enumerate()
            {
                let store_indices = snapshot.store_indices(binding.store_item_id());
                let (store_part, status) = match binding.checksum_status() {
                    ChecksumStatus::Absent => (None, PackageChecksumStatus::Absent),
                    ChecksumStatus::Malformed(value) => {
                        (None, PackageChecksumStatus::Malformed(value))
                    },
                    ChecksumStatus::Unchecked(expected) => {
                        if store_indices.is_empty() {
                            (None, PackageChecksumStatus::MissingStore)
                        } else if store_indices.len() > 1 {
                            return Err(invalid(format!(
                                "Custom XML itemID '{}' is ambiguous",
                                binding.store_item_id()
                            )));
                        } else {
                            let store = snapshot
                                .stores
                                .get(store_indices[0])
                                .ok_or_else(|| invalid("Custom XML GUID index is corrupt"))?;
                            let actual = if let Some(value) = memo.get(&store.part) {
                                value.clone()
                            } else {
                                if memo.len() >= snapshot.limits.max_crc_parts {
                                    return Err(invalid("unique CRC part limit exceeded"));
                                }
                                total = total
                                    .checked_add(store.data.len())
                                    .ok_or_else(|| invalid("aggregate CRC byte count overflow"))?;
                                if total > snapshot.limits.max_crc_bytes {
                                    return Err(invalid("aggregate CRC byte limit exceeded"));
                                }
                                let value =
                                    Checksum::compute(&store.data, &snapshot.limits.controls)?;
                                memo.try_reserve(1)
                                    .map_err(alloc("content-control CRC memo"))?;
                                memo.insert(store.part.clone(), value.clone());
                                value
                            };
                            let status = if expected.as_bytes() == actual.as_bytes() {
                                PackageChecksumStatus::Matches
                            } else {
                                PackageChecksumStatus::Mismatch { expected, actual }
                            };
                            (Some(store.part.clone()), status)
                        }
                    },
                    ChecksumStatus::Matches | ChecksumStatus::Mismatch { .. } => {
                        return Err(invalid("unexpected pre-verified checksum state"));
                    },
                };
                if require_matches == Some(true)
                    && !matches!(
                        status,
                        PackageChecksumStatus::Absent | PackageChecksumStatus::Matches
                    )
                {
                    return Err(invalid(
                        "changed content-control publication has an unverifiable checksum",
                    ));
                }
                entries.push(ChecksumEntry {
                    part: story.part.clone(),
                    occurrence: occurrence.ordinal(),
                    binding: binding_index,
                    flavor: binding.flavor(),
                    control_id: occurrence.id(),
                    store_item_id: binding.store_item_id().to_owned(),
                    store_part,
                    status,
                });
            }
        }
    }
    Ok(entries)
}

fn capture_stores(
    package: &OpcPackage,
    limits: &PackageLimits,
) -> Result<(Arc<[Store]>, Arc<[u8]>)> {
    preflight_custom_xml_graph(package, limits)?;
    let items = custom_xml::discover(package)?;
    let mut graph_bytes = b"litchi.docx.sdt.custom-xml.v1\0".len();
    charge_custom_graph_bytes(&mut graph_bytes, 8, limits.max_custom_graph_bytes)?;
    for item in &items {
        charge_custom_graph_bytes(&mut graph_bytes, 1, limits.max_custom_graph_bytes)?;
        charge_custom_graph_field(
            &mut graph_bytes,
            item.source().as_str().len(),
            limits.max_custom_graph_bytes,
        )?;
        charge_custom_graph_field(
            &mut graph_bytes,
            item.rel_id().len(),
            limits.max_custom_graph_bytes,
        )?;
        charge_custom_graph_field(
            &mut graph_bytes,
            item.source_relationship().relationship_type.len(),
            limits.max_custom_graph_bytes,
        )?;
        charge_custom_graph_field(
            &mut graph_bytes,
            item.source_relationship().target.len(),
            limits.max_custom_graph_bytes,
        )?;
        charge_custom_graph_bytes(&mut graph_bytes, 1, limits.max_custom_graph_bytes)?;
        charge_custom_graph_field(
            &mut graph_bytes,
            item.part().as_str().len(),
            limits.max_custom_graph_bytes,
        )?;
        charge_custom_graph_field(
            &mut graph_bytes,
            item.content_type().len(),
            limits.max_custom_graph_bytes,
        )?;
        let props_xml = item
            .props_part()
            .map(|part| package.get_part(part).map(|value| value.blob()))
            .transpose()?;
        charge_custom_graph_field(
            &mut graph_bytes,
            props_xml.map_or(0, <[u8]>::len),
            limits.max_custom_graph_bytes,
        )?;
        charge_custom_graph_field(
            &mut graph_bytes,
            item.props().map_or(0, |props| props.id.len()),
            limits.max_custom_graph_bytes,
        )?;
        charge_custom_graph_bytes(&mut graph_bytes, 8, limits.max_custom_graph_bytes)?;
        if let Some(props) = item.props() {
            for schema in &props.schemas {
                charge_custom_graph_field(
                    &mut graph_bytes,
                    schema.len(),
                    limits.max_custom_graph_bytes,
                )?;
            }
        }
        charge_custom_graph_bytes(&mut graph_bytes, 8, limits.max_custom_graph_bytes)?;
        for relationship in item.relationships() {
            charge_custom_graph_bytes(&mut graph_bytes, 2, limits.max_custom_graph_bytes)?;
            charge_custom_graph_field(
                &mut graph_bytes,
                relationship.id.len(),
                limits.max_custom_graph_bytes,
            )?;
            charge_custom_graph_field(
                &mut graph_bytes,
                relationship.relationship_type.len(),
                limits.max_custom_graph_bytes,
            )?;
            charge_custom_graph_field(
                &mut graph_bytes,
                relationship.target.len(),
                limits.max_custom_graph_bytes,
            )?;
        }
    }
    let mut by_part = HashMap::<PackURI, Store>::new();
    let mut graph = Vec::new();
    graph
        .try_reserve_exact(graph_bytes)
        .map_err(alloc("Custom XML graph token"))?;
    graph.extend_from_slice(b"litchi.docx.sdt.custom-xml.v1\0");
    put_number(&mut graph, items.len())?;
    for item in &items {
        graph.push(1);
        put_field(&mut graph, item.source().as_str().as_bytes())?;
        put_field(&mut graph, item.rel_id().as_bytes())?;
        put_field(
            &mut graph,
            item.source_relationship().relationship_type.as_bytes(),
        )?;
        put_field(&mut graph, item.source_relationship().target.as_bytes())?;
        graph.push(u8::from(item.source_relationship().is_external()));
        put_field(&mut graph, item.part().as_str().as_bytes())?;
        put_field(&mut graph, item.content_type().as_bytes())?;
        let props_xml = item
            .props_part()
            .map(|part| package.get_part(part).map(|value| value.blob()))
            .transpose()?;
        put_field(&mut graph, props_xml.unwrap_or_default())?;
        let id = item
            .props()
            .map(|props| props.id.clone())
            .unwrap_or_default();
        put_field(&mut graph, id.as_bytes())?;
        put_number(
            &mut graph,
            item.props().map_or(0, |props| props.schemas.len()),
        )?;
        if let Some(props) = item.props() {
            for schema in &props.schemas {
                put_field(&mut graph, schema.as_bytes())?;
            }
        }
        put_number(&mut graph, item.relationships().len())?;
        for relationship in item.relationships() {
            graph.push(2);
            put_field(&mut graph, relationship.id.as_bytes())?;
            put_field(&mut graph, relationship.relationship_type.as_bytes())?;
            put_field(&mut graph, relationship.target.as_bytes())?;
            graph.push(u8::from(relationship.is_external()));
        }
        if graph.len() > limits.max_custom_graph_bytes {
            return Err(invalid(
                "Custom XML graph exceeds configured stale-check limit",
            ));
        }
        if !by_part.contains_key(item.part()) {
            by_part
                .try_reserve(1)
                .map_err(alloc("Custom XML store index"))?;
            by_part.insert(
                item.part().clone(),
                Store {
                    id,
                    part: item.part().clone(),
                    content_type: item.content_type().to_owned(),
                    data: package.get_part(item.part())?.blob_arc(),
                    props_part: item.props_part().cloned(),
                },
            );
        }
    }
    let mut stores = by_part.into_values().collect::<Vec<_>>();
    stores.sort_unstable_by(|left, right| left.part.as_str().cmp(right.part.as_str()));
    Ok((
        Arc::from(stores.into_boxed_slice()),
        Arc::from(graph.into_boxed_slice()),
    ))
}

fn preflight_custom_xml_graph(package: &OpcPackage, limits: &PackageLimits) -> Result<()> {
    const TRANSITIONAL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
    const STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";
    const TRANSITIONAL_PROPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
    const STRICT_PROPS: &str =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/customXmlProps";

    let mut parts = HashSet::<PackURI>::new();
    let mut payload_bytes = 0usize;
    let mut graph_bytes = b"litchi.docx.sdt.custom-xml.v1\0".len();
    charge_custom_graph_bytes(&mut graph_bytes, 8, limits.max_custom_graph_bytes)?;
    let mut relationship_count = 0usize;
    if package
        .rels()
        .iter()
        .any(|relationship| matches!(relationship.reltype(), TRANSITIONAL | STRICT))
    {
        return Err(invalid(
            "package root cannot source a Custom XML Data Storage relationship",
        ));
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if relationship.reltype() != TRANSITIONAL && relationship.reltype() != STRICT {
                continue;
            }
            relationship_count = relationship_count
                .checked_add(1)
                .ok_or_else(|| invalid("Custom XML relationship count overflow"))?;
            if relationship_count > limits.max_custom_relationships {
                return Err(invalid(
                    "Custom XML relationship count exceeds configured limit",
                ));
            }
            if relationship.is_external() {
                return Err(invalid("Custom XML data relationship must be internal"));
            }
            charge_custom_graph_bytes(&mut graph_bytes, 2, limits.max_custom_graph_bytes)?;
            for length in [
                source.partname().as_str().len(),
                relationship.r_id().len(),
                relationship.reltype().len(),
                relationship.target_ref().len(),
            ] {
                charge_custom_graph_field(&mut graph_bytes, length, limits.max_custom_graph_bytes)?;
            }
            let part = relationship
                .target_partname()
                .map_err(|error| invalid(format!("invalid Custom XML target: {error}")))?;
            let data = package.get_part(&part)?;
            for length in [part.as_str().len(), data.content_type().len()] {
                charge_custom_graph_field(&mut graph_bytes, length, limits.max_custom_graph_bytes)?;
            }
            if !parts.contains(&part) {
                if parts.len() >= limits.max_crc_parts {
                    return Err(invalid("unique Custom XML payload limit exceeded"));
                }
                payload_bytes = payload_bytes
                    .checked_add(data.blob().len())
                    .ok_or_else(|| invalid("aggregate Custom XML payload size overflow"))?;
                if payload_bytes > limits.max_crc_bytes {
                    return Err(invalid(
                        "aggregate Custom XML payload bytes exceed configured limit",
                    ));
                }
                parts
                    .try_reserve(1)
                    .map_err(alloc("Custom XML payload preflight index"))?;
                parts.insert(part.clone());
            }
            charge_custom_graph_bytes(&mut graph_bytes, 8, limits.max_custom_graph_bytes)?;
            for child in data.rels().iter() {
                relationship_count = relationship_count
                    .checked_add(1)
                    .ok_or_else(|| invalid("Custom XML relationship count overflow"))?;
                if relationship_count > limits.max_custom_relationships {
                    return Err(invalid(
                        "Custom XML relationship count exceeds configured limit",
                    ));
                }
                charge_custom_graph_bytes(&mut graph_bytes, 2, limits.max_custom_graph_bytes)?;
                for length in [
                    child.r_id().len(),
                    child.reltype().len(),
                    child.target_ref().len(),
                ] {
                    charge_custom_graph_field(
                        &mut graph_bytes,
                        length,
                        limits.max_custom_graph_bytes,
                    )?;
                }
                if !matches!(child.reltype(), TRANSITIONAL_PROPS | STRICT_PROPS) {
                    continue;
                }
                if child.is_external() {
                    return Err(invalid(
                        "Custom XML properties relationship must be internal",
                    ));
                }
                let props_name = child.target_partname().map_err(|error| {
                    invalid(format!("invalid Custom XML properties target: {error}"))
                })?;
                let props = package.get_part(&props_name)?;
                charge_custom_graph_field(
                    &mut graph_bytes,
                    props.blob().len(),
                    limits.max_custom_graph_bytes,
                )?;
                // The common discovery layer parses and clones the bounded
                // semantic ID/schema metadata in addition to retaining the raw
                // properties bytes. Charge a second raw-size envelope before it.
                charge_custom_graph_field(
                    &mut graph_bytes,
                    props.blob().len(),
                    limits.max_custom_graph_bytes,
                )?;
            }
        }
    }
    Ok(())
}

fn charge_custom_graph_field(charged: &mut usize, length: usize, limit: usize) -> Result<()> {
    let amount = 8usize
        .checked_add(length)
        .ok_or_else(|| invalid("Custom XML graph field size overflow"))?;
    charge_custom_graph_bytes(charged, amount, limit)
}

fn charge_custom_graph_bytes(charged: &mut usize, amount: usize, limit: usize) -> Result<()> {
    *charged = charged
        .checked_add(amount)
        .ok_or_else(|| invalid("Custom XML graph size overflow"))?;
    if *charged > limit {
        return Err(invalid(
            "Custom XML graph exceeds configured stale-check limit",
        ));
    }
    Ok(())
}

fn index_stores(stores: &[Store]) -> Result<Arc<HashMap<String, Vec<usize>>>> {
    let mut index = HashMap::<String, Vec<usize>>::new();
    index
        .try_reserve(stores.len())
        .map_err(alloc("Custom XML GUID index"))?;
    for (position, store) in stores.iter().enumerate() {
        if store.id.is_empty() {
            continue;
        }
        let positions = index.entry(store.id.to_ascii_lowercase()).or_default();
        positions
            .try_reserve(1)
            .map_err(alloc("Custom XML GUID occurrences"))?;
        positions.push(position);
    }
    Ok(Arc::new(index))
}

fn signature_token(package: &OpcPackage, limit: usize) -> Result<Arc<[u8]>> {
    // Account for the complete graph while all payloads are still borrowed.
    // In particular, never copy an attacker-controlled signature blob and
    // only then discover that it exceeds the configured stale-check budget.
    let mut charged = SIGNATURE_TOKEN_MAGIC
        .len()
        .checked_add(16)
        .ok_or_else(|| invalid("signature token size overflow"))?;
    if charged > limit {
        return Err(invalid(
            "signature graph exceeds configured stale-check limit",
        ));
    }
    let mut relationship_count = 0usize;
    let mut part_names = Vec::<PackURI>::new();
    let mut seen = HashSet::<PackURI>::new();
    for relationship in package.rels().iter() {
        if !root_signature_relationship(package, relationship) {
            continue;
        }
        relationship_count = relationship_count
            .checked_add(1)
            .ok_or_else(|| invalid("signature relationship count overflow"))?;
        charge_signature_bytes(&mut charged, 2, limit)?;
        charge_signature_field(&mut charged, relationship.r_id().len(), limit)?;
        charge_signature_field(&mut charged, relationship.reltype().len(), limit)?;
        charge_signature_field(&mut charged, relationship.target_ref().len(), limit)?;
        if !relationship.is_external() {
            let target = relationship
                .target_partname()
                .map_err(|error| invalid(format!("invalid signature target: {error}")))?;
            add_signature_part(
                package,
                target,
                &mut seen,
                &mut part_names,
                &mut charged,
                limit,
            )?;
        }
    }
    for part in package.iter_parts() {
        if is_signature_part(part) {
            add_signature_part(
                package,
                part.partname().clone(),
                &mut seen,
                &mut part_names,
                &mut charged,
                limit,
            )?;
        }
    }
    let mut cursor = 0usize;
    while cursor < part_names.len() {
        let source = part_names[cursor].clone();
        cursor += 1;
        let part = package.get_part(&source)?;
        for relationship in part.rels().iter() {
            if relationship.is_external()
                || !is_signature_relationship(relationship.reltype(), relationship.target_ref())
            {
                continue;
            }
            let target = relationship
                .target_partname()
                .map_err(|error| invalid(format!("invalid signature target: {error}")))?;
            add_signature_part(
                package,
                target,
                &mut seen,
                &mut part_names,
                &mut charged,
                limit,
            )?;
        }
    }
    part_names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));

    let mut token = Vec::new();
    token
        .try_reserve_exact(charged)
        .map_err(alloc("signature graph token"))?;
    token.extend_from_slice(SIGNATURE_TOKEN_MAGIC);
    put_number(&mut token, relationship_count)?;
    for relationship in package.rels().iter() {
        if !root_signature_relationship(package, relationship) {
            continue;
        }
        token.push(1);
        put_field(&mut token, relationship.r_id().as_bytes())?;
        put_field(&mut token, relationship.reltype().as_bytes())?;
        put_field(&mut token, relationship.target_ref().as_bytes())?;
        token.push(u8::from(relationship.is_external()));
    }
    put_number(&mut token, part_names.len())?;
    for part_name in part_names {
        let part = package.get_part(&part_name)?;
        token.push(2);
        put_field(&mut token, part.partname().as_str().as_bytes())?;
        put_field(&mut token, part.content_type().as_bytes())?;
        put_field(&mut token, part.blob())?;
        put_number(&mut token, part.rels().len())?;
        for relationship in part.rels().iter() {
            token.push(3);
            put_field(&mut token, relationship.r_id().as_bytes())?;
            put_field(&mut token, relationship.reltype().as_bytes())?;
            put_field(&mut token, relationship.target_ref().as_bytes())?;
            token.push(u8::from(relationship.is_external()));
        }
    }
    debug_assert_eq!(token.len(), charged);
    Ok(Arc::from(token.into_boxed_slice()))
}

fn add_signature_part(
    package: &OpcPackage,
    part_name: PackURI,
    seen: &mut HashSet<PackURI>,
    part_names: &mut Vec<PackURI>,
    charged: &mut usize,
    limit: usize,
) -> Result<()> {
    if seen.contains(&part_name) {
        return Ok(());
    }
    let part = package.get_part(&part_name)?;
    charge_signature_bytes(charged, 1, limit)?;
    charge_signature_field(charged, part.partname().as_str().len(), limit)?;
    charge_signature_field(charged, part.content_type().len(), limit)?;
    charge_signature_field(charged, part.blob().len(), limit)?;
    charge_signature_bytes(charged, 8, limit)?;
    for relationship in part.rels().iter() {
        charge_signature_bytes(charged, 2, limit)?;
        charge_signature_field(charged, relationship.r_id().len(), limit)?;
        charge_signature_field(charged, relationship.reltype().len(), limit)?;
        charge_signature_field(charged, relationship.target_ref().len(), limit)?;
    }
    seen.try_reserve(1).map_err(alloc("signature part index"))?;
    part_names
        .try_reserve(1)
        .map_err(alloc("signature parts"))?;
    seen.insert(part_name.clone());
    part_names.push(part_name);
    Ok(())
}

fn root_signature_relationship(
    package: &OpcPackage,
    relationship: &litchi_opc::Relationship,
) -> bool {
    if is_signature_relationship(relationship.reltype(), relationship.target_ref()) {
        return true;
    }
    if relationship.is_external() {
        return false;
    }
    relationship
        .target_partname()
        .ok()
        .and_then(|part| package.get_part(&part).ok())
        .is_some_and(is_signature_part)
}

fn is_signature_part(part: &dyn litchi_opc::Part) -> bool {
    starts_with_ascii_case_insensitive(part.partname().as_str(), "/_xmlsignatures/")
        || contains_ascii_case_insensitive(part.content_type(), "digital-signature")
}

fn charge_signature_field(charged: &mut usize, length: usize, limit: usize) -> Result<()> {
    let amount = 8usize
        .checked_add(length)
        .ok_or_else(|| invalid("signature field size overflow"))?;
    charge_signature_bytes(charged, amount, limit)
}

fn charge_signature_bytes(charged: &mut usize, amount: usize, limit: usize) -> Result<()> {
    *charged = charged
        .checked_add(amount)
        .ok_or_else(|| invalid("signature token size overflow"))?;
    if *charged > limit {
        return Err(invalid(
            "signature graph exceeds configured stale-check limit",
        ));
    }
    Ok(())
}

fn is_signature_relationship(reltype: &str, target: &str) -> bool {
    contains_ascii_case_insensitive(reltype, "digital-signature")
        || contains_ascii_case_insensitive(target, "_xmlsignatures")
}

fn starts_with_ascii_case_insensitive(value: &str, prefix: &str) -> bool {
    value
        .as_bytes()
        .get(..prefix.len())
        .is_some_and(|candidate| candidate.eq_ignore_ascii_case(prefix.as_bytes()))
}

fn contains_ascii_case_insensitive(value: &str, needle: &str) -> bool {
    value
        .as_bytes()
        .windows(needle.len())
        .any(|candidate| candidate.eq_ignore_ascii_case(needle.as_bytes()))
}

fn unsigned_signature_token() -> Result<Arc<[u8]>> {
    let capacity = SIGNATURE_TOKEN_MAGIC
        .len()
        .checked_add(16)
        .ok_or_else(|| invalid("unsigned signature token size overflow"))?;
    let mut token = Vec::new();
    token
        .try_reserve_exact(capacity)
        .map_err(alloc("unsigned signature token"))?;
    token.extend_from_slice(SIGNATURE_TOKEN_MAGIC);
    put_number(&mut token, 0)?;
    put_number(&mut token, 0)?;
    Ok(Arc::from(token.into_boxed_slice()))
}

fn is_unsigned_signature_token(token: &[u8]) -> bool {
    token.len() == SIGNATURE_TOKEN_MAGIC.len() + 16
        && token.starts_with(SIGNATURE_TOKEN_MAGIC)
        && token[SIGNATURE_TOKEN_MAGIC.len()..]
            .iter()
            .all(|byte| *byte == 0)
}

fn require_current(
    expected: &PackageSnapshot,
    current: &PackageSnapshot,
    signatures: bool,
) -> Result<()> {
    if expected.topology != current.topology {
        return Err(invalid("content-control story topology is stale"));
    }
    if signatures && expected.signatures != current.signatures {
        return Err(invalid("package signature topology is stale"));
    }
    if expected.custom_graph != current.custom_graph
        || !same_stores(&expected.stores, &current.stores)
    {
        return Err(invalid("Custom XML checksum source is stale"));
    }
    if expected.stories.len() != current.stories.len()
        || expected
            .stories
            .iter()
            .zip(current.stories.iter())
            .any(|(left, right)| {
                left.part != right.part || left.snapshot.source() != right.snapshot.source()
            })
    {
        return Err(invalid("content-control story source is stale"));
    }
    Ok(())
}

fn same_stores(left: &[Store], right: &[Store]) -> bool {
    left.len() == right.len()
        && left.iter().zip(right).all(|(left, right)| {
            left.id == right.id
                && left.part == right.part
                && left.content_type == right.content_type
                && left.data.as_slice() == right.data.as_slice()
                && left.props_part == right.props_part
        })
}

fn put_field(output: &mut Vec<u8>, value: &[u8]) -> Result<()> {
    let length = u64::try_from(value.len())
        .map_err(|_| invalid("content-control topology field is too large"))?;
    let additional = 8usize
        .checked_add(value.len())
        .ok_or_else(|| invalid("content-control topology size overflow"))?;
    output
        .try_reserve(additional)
        .map_err(alloc("content-control topology token"))?;
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(value);
    Ok(())
}

fn put_number(output: &mut Vec<u8>, value: usize) -> Result<()> {
    let value =
        u64::try_from(value).map_err(|_| invalid("content-control topology count is too large"))?;
    output
        .try_reserve(8)
        .map_err(alloc("content-control topology token"))?;
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn alloc(resource: &'static str) -> impl FnOnce(std::collections::TryReserveError) -> Error {
    move |source| Error::Allocation { resource, source }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_ooxml_common::custom_xml::Conformance;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::part::BlobPart;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const HASH: &str = "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash";
    const ITEM: &str = "{11111111-1111-4111-8111-111111111111}";
    const ITEM_TWO: &str = "{22222222-2222-4222-8222-222222222222}";

    fn replace_main(package: &mut Package, xml: Vec<u8>) -> PackURI {
        let main = package
            .opc_package()
            .main_document_part()
            .unwrap()
            .partname()
            .clone();
        package
            .edit_opc(|opc| {
                opc.get_part_mut(&main)?.set_blob(xml);
                Ok::<_, Error>(())
            })
            .unwrap();
        main
    }

    fn mark_signed(package: &mut Package) {
        package
            .edit_opc(|opc| {
                opc.try_add_part(Box::new(BlobPart::new(
                    PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                    ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                    Vec::new(),
                )))?;
                opc.rels_mut().add_relationship(
                    rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                    "_xmlsignatures/origin.sigs".to_owned(),
                    "rSignature".to_owned(),
                    false,
                );
                Ok::<_, Error>(())
            })
            .unwrap();
    }

    #[test]
    fn package_entry_points_refuse_pending_facade_state_before_claiming_patch() {
        let mut package = Package::new().unwrap();
        let snapshot = package.content_control_snapshot().unwrap();
        let commit = snapshot.edit().unwrap().commit().unwrap();
        assert!(!commit.changed());
        assert!(!commit.patch().is_applied());

        package
            .document_mut()
            .unwrap()
            .add_paragraph_with_text("pending facade state");

        assert!(package.content_control_snapshot().is_err());
        assert!(
            package
                .content_control_snapshot_with_limits(PackageLimits::default())
                .is_err()
        );
        assert!(package.verify_content_control_checksums().is_err());
        assert!(package.apply_content_controls(&commit).is_err());
        assert!(!commit.patch().is_applied());
        assert!(package.apply_content_control_patch(commit.patch()).is_err());
        assert!(!commit.patch().is_applied());
    }

    #[test]
    fn clean_exact_noop_still_bypasses_opc_publication() {
        let mut package = Package::new().unwrap();
        let main = package.opc_package().main_document_part().unwrap();
        let before = main.blob_arc();
        let snapshot = package.content_control_snapshot().unwrap();
        let commit = snapshot.edit().unwrap().commit().unwrap();

        package.apply_content_controls(&commit).unwrap();

        let after = package
            .opc_package()
            .main_document_part()
            .unwrap()
            .blob_arc();
        assert!(Arc::ptr_eq(&before, &after));
        assert!(commit.patch().is_applied());
    }

    #[test]
    fn changed_signed_patch_requires_explicit_unsign_but_noop_preserves_signature() {
        let mut package = Package::new().unwrap();
        let main = replace_main(
            &mut package,
            format!(
                r#"<w:document xmlns:w="{W}"><w:body><w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent/></w:sdt></w:body></w:document>"#,
            )
            .into_bytes(),
        );
        mark_signed(&mut package);

        let signed = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
        let snapshot = package.content_control_snapshot().unwrap();
        let noop = snapshot.edit().unwrap().commit().unwrap();
        package.apply_content_controls(&noop).unwrap();
        assert!(package.opc_package().contains_part(&signed));

        let snapshot = package.content_control_snapshot().unwrap();
        let mut transaction = snapshot.edit().unwrap();
        transaction
            .set_formatting_allowed(&main, 0, Some(super::super::FormattingAllowed::Allowed))
            .unwrap();
        let commit = transaction.commit().unwrap();
        let inverse = commit.patch().inverse().unwrap();
        assert!(package.apply_content_controls(&commit).is_err());
        assert!(!commit.patch().is_applied());
        assert!(package.opc_package().contains_part(&signed));

        package.unsign();
        assert!(!package.opc_package().contains_part(&signed));
        package.apply_content_controls(&commit).unwrap();
        assert!(commit.patch().is_applied());
        let redo = inverse.inverse().unwrap();
        package.apply_content_control_patch(&inverse).unwrap();
        package.apply_content_control_patch(&redo).unwrap();
        assert!(
            std::str::from_utf8(package.opc_package().get_part(&main).unwrap().blob())
                .unwrap()
                .contains("formattingAllowed=\"1\"")
        );
    }

    #[test]
    fn signature_inventory_follows_relationships_and_content_types_not_part_names() {
        let mut relationship_owned = Package::new().unwrap();
        relationship_owned
            .edit_opc(|opc| {
                opc.try_add_part(Box::new(BlobPart::new(
                    PackURI::new("/security/arbitrary-origin.bin").unwrap(),
                    "application/octet-stream".to_owned(),
                    b"relationship-owned".to_vec(),
                )))?;
                opc.rels_mut().add_relationship(
                    rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                    "security/arbitrary-origin.bin".to_owned(),
                    "rArbitrarySignature".to_owned(),
                    false,
                );
                Ok::<_, Error>(())
            })
            .unwrap();
        let snapshot = relationship_owned.content_control_snapshot().unwrap();
        assert!(!is_unsigned_signature_token(&snapshot.signatures));

        let mut content_typed = Package::new().unwrap();
        content_typed
            .edit_opc(|opc| {
                opc.try_add_part(Box::new(BlobPart::new(
                    PackURI::new("/security/content-typed.bin").unwrap(),
                    ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                    b"content-type-owned".to_vec(),
                )))?;
                Ok::<_, Error>(())
            })
            .unwrap();
        let snapshot = content_typed.content_control_snapshot().unwrap();
        assert!(!is_unsigned_signature_token(&snapshot.signatures));

        let bounded = PackageLimits {
            max_signature_bytes: SIGNATURE_TOKEN_MAGIC.len() + 16,
            ..PackageLimits::default()
        };
        assert!(
            content_typed
                .content_control_snapshot_with_limits(bounded)
                .is_err()
        );
    }

    #[test]
    fn changed_patch_inverse_accepts_unsigned_state_and_refuses_new_signature() {
        let mut package = Package::new().unwrap();
        let main = replace_main(
            &mut package,
            format!(
                r#"<w:document xmlns:w="{W}"><w:body><w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent/></w:sdt></w:body></w:document>"#,
            )
            .into_bytes(),
        );
        let original = package
            .opc_package()
            .get_part(&main)
            .unwrap()
            .blob()
            .to_vec();

        let snapshot = package.content_control_snapshot().unwrap();
        let mut transaction = snapshot.edit().unwrap();
        transaction
            .set_formatting_allowed(&main, 0, Some(super::super::FormattingAllowed::Allowed))
            .unwrap();
        let commit = transaction.commit().unwrap();
        let inverse = commit.patch().inverse().unwrap();
        package.apply_content_controls(&commit).unwrap();
        package.apply_content_control_patch(&inverse).unwrap();
        assert_eq!(
            package.opc_package().get_part(&main).unwrap().blob(),
            original
        );

        let snapshot = package.content_control_snapshot().unwrap();
        let mut transaction = snapshot.edit().unwrap();
        transaction
            .set_formatting_allowed(&main, 0, Some(super::super::FormattingAllowed::Allowed))
            .unwrap();
        let commit = transaction.commit().unwrap();
        let inverse = commit.patch().inverse().unwrap();
        package.apply_content_controls(&commit).unwrap();
        mark_signed(&mut package);

        assert!(package.apply_content_control_patch(&inverse).is_err());
        assert!(!inverse.is_applied());
    }

    #[test]
    fn bulk_refresh_quota_failure_is_atomic_and_retryable() {
        let payload = b"<root><value>bounded</value></root>";
        let checksum = Checksum::compute(b"<stale/>", &Limits::default())
            .unwrap()
            .to_base64();
        let mut package = Package::new().unwrap();
        package
            .add_custom_xml(crate::custom_xml::NewStore {
                xml: payload.to_vec(),
                content_type: "application/xml".to_owned(),
                id: ITEM.to_owned(),
                schemas: Vec::new(),
                conformance: Conformance::Transitional,
            })
            .unwrap();
        let controls = format!(
            r#"<w:sdt><w:sdtPr><w:dataBinding w:xpath="/root" w:storeItemID="{ITEM}" h:storeItemChecksum="{checksum}"/></w:sdtPr><w:sdtContent/></w:sdt><w:sdt><w:sdtPr><w:dataBinding w:xpath="/root/value" w:storeItemID="{ITEM}" h:storeItemChecksum="{checksum}"/></w:sdtPr><w:sdtContent/></w:sdt>"#,
        );
        let main = replace_main(
            &mut package,
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:h="{HASH}" mc:Ignorable="h"><w:body>{controls}</w:body></w:document>"#,
            )
            .into_bytes(),
        );
        let limits = PackageLimits {
            max_mutations: 1,
            ..PackageLimits::default()
        };
        let snapshot = package
            .content_control_snapshot_with_limits(limits)
            .unwrap();
        let mut transaction = snapshot.edit().unwrap();

        assert!(transaction.refresh_checksums().is_err());
        assert_eq!(
            transaction
                .transactions
                .iter()
                .map(|story| story.edits().count())
                .sum::<usize>(),
            0
        );
        assert!(transaction.refresh_checksums().is_err());
        assert_eq!(
            transaction
                .transactions
                .iter()
                .map(|story| story.edits().count())
                .sum::<usize>(),
            0
        );

        transaction.refresh_checksum(&main, 0).unwrap();
        transaction.refresh_checksum(&main, 0).unwrap();
        assert_eq!(
            transaction
                .transactions
                .iter()
                .map(|story| story.edits().count())
                .sum::<usize>(),
            1
        );
        assert!(transaction.refresh_checksums().is_err());
        assert_eq!(
            transaction
                .transactions
                .iter()
                .map(|story| story.edits().count())
                .sum::<usize>(),
            1
        );
    }

    #[test]
    fn duplicate_store_guid_is_rejected_before_checksum_reporting() {
        let mut package = Package::new().unwrap();
        for id in [ITEM, ITEM_TWO] {
            package
                .add_custom_xml(crate::custom_xml::NewStore {
                    xml: format!("<root id=\"{id}\"/>").into_bytes(),
                    content_type: "application/xml".to_owned(),
                    id: id.to_owned(),
                    schemas: Vec::new(),
                    conformance: Conformance::Transitional,
                })
                .unwrap();
        }
        let first = PackURI::new("/customXml/itemProps1.xml").unwrap();
        let second = PackURI::new("/customXml/itemProps2.xml").unwrap();
        package
            .edit_opc(|opc| {
                let duplicate = opc.get_part(&first)?.blob().to_vec();
                opc.get_part_mut(&second)?.set_blob(duplicate);
                Ok::<_, Error>(())
            })
            .unwrap();

        assert!(package.content_control_snapshot().is_err());
        assert!(package.verify_content_control_checksums().is_err());
    }

    #[test]
    fn package_binding_and_report_limits_are_independent() {
        let mut package = Package::new().unwrap();
        replace_main(
            &mut package,
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w15"><w:body><w:sdtPr><w:dataBinding w:xpath="/one" w:storeItemID="{ITEM}"/><w15:dataBinding w:xpath="/two" w:storeItemID="{ITEM_TWO}"/></w:sdtPr></w:body></w:document>"#,
            )
            .into_bytes(),
        );

        let binding_limited = PackageLimits {
            max_bindings: 1,
            ..PackageLimits::default()
        };
        assert!(
            package
                .content_control_snapshot_with_limits(binding_limited)
                .is_err()
        );

        let report_limited = PackageLimits {
            max_report_entries: 1,
            ..PackageLimits::default()
        };
        let snapshot = package
            .content_control_snapshot_with_limits(report_limited)
            .unwrap();
        assert!(snapshot.verify_checksums().is_err());
    }

    #[test]
    fn custom_xml_graph_is_bounded_before_discovery() {
        let mut package = Package::new().unwrap();
        package
            .add_custom_xml(crate::custom_xml::NewStore {
                xml: b"<root/>".to_vec(),
                content_type: "application/xml".to_owned(),
                id: ITEM.to_owned(),
                schemas: Vec::new(),
                conformance: Conformance::Transitional,
            })
            .unwrap();

        let relationship_limited = PackageLimits {
            max_custom_relationships: 1,
            ..PackageLimits::default()
        };
        assert!(
            package
                .content_control_snapshot_with_limits(relationship_limited)
                .is_err()
        );

        let props_limited = PackageLimits {
            max_custom_graph_bytes: 512,
            ..PackageLimits::default()
        };
        assert!(
            package
                .content_control_snapshot_with_limits(props_limited)
                .is_err()
        );
    }

    #[test]
    fn package_refresh_edits_only_selected_alternate_content_binding() {
        let payload = b"<root><value>selected</value></root>";
        let stale = Checksum::compute(b"<stale/>", &Limits::default())
            .unwrap()
            .to_base64();
        let expected = Checksum::compute(payload, &Limits::default())
            .unwrap()
            .to_base64();
        let mut package = Package::new().unwrap();
        package
            .add_custom_xml(crate::custom_xml::NewStore {
                xml: payload.to_vec(),
                content_type: "application/xml".to_owned(),
                id: ITEM.to_owned(),
                schemas: Vec::new(),
                conformance: Conformance::Transitional,
            })
            .unwrap();
        let main = replace_main(
            &mut package,
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:h="{HASH}" mc:Ignorable="w15 h"><w:body><w:sdtPr><mc:AlternateContent><mc:Choice Requires="w15"><w:dataBinding w:xpath="/active" w:storeItemID="{ITEM}" h:storeItemChecksum="{stale}"/></mc:Choice><mc:Fallback><w:dataBinding w:xpath="/inactive" w:storeItemID="{ITEM}" h:storeItemChecksum="{stale}"/></mc:Fallback></mc:AlternateContent></w:sdtPr></w:body></w:document>"#,
            )
            .into_bytes(),
        );

        let snapshot = package.content_control_snapshot().unwrap();
        let mut transaction = snapshot.edit().unwrap();
        transaction.refresh_checksums().unwrap();
        let commit = transaction.commit().unwrap();
        package.apply_content_controls(&commit).unwrap();
        let output =
            std::str::from_utf8(package.opc_package().get_part(&main).unwrap().blob()).unwrap();
        assert_eq!(output.matches(&expected).count(), 1);
        assert_eq!(output.matches(&stale).count(), 1);
        assert!(output.find(&expected).unwrap() < output.find("/inactive").unwrap());
    }
}
