#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by the revision owner and publication lifecycle"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are deliberately refined after XML validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local XML names mirror the corresponding WordprocessingML roles"
)]
//! Source-preserving authoring and disposition of ordinary WordprocessingML
//! tracked changes.
//!
//! This owner deliberately has a small boundary. It resolves inline
//! `w:ins`/`w:del` changes and complete `w:moveFrom`/`w:moveTo` pairs,
//! including their paired move range markers. Paragraph/property changes,
//! table-row changes, and markup-compatibility alternatives are refused until
//! they have a separate typed owner. Revisions contained by fields or
//! structured document controls are refused because this owner never evaluates
//! document fields, macros, controls, or embedded content.

use std::{
    collections::{HashMap, HashSet},
    sync::{
        Arc,
        atomic::{AtomicU8, Ordering},
    },
};

use quick_xml::{
    XmlVersion,
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use crate::package::story::{StoryLimits, capture};
use crate::{Error, Package, Result};
use litchi_ooxml_common::properties::time::DateTime;
use litchi_opc::PackURI;

const W_TRANSITIONAL: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const MC: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML: &[u8] = b"http://www.w3.org/XML/1998/namespace";

const MAX_SOURCE_BYTES: usize = 256 * 1024 * 1024;
const MAX_EVENTS: usize = 16_000_000;
const MAX_DEPTH: usize = 128;
const MAX_CHANGES: usize = 1_000_000;
const MAX_RANGES: usize = 1_000_000;
const MAX_METADATA_BYTES: usize = 64 * 1024 * 1024;
const MAX_ATTRIBUTE_BYTES: usize = 1024 * 1024;
const MAX_ATTRIBUTES: usize = 4_096;
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;

const GATE_READY: u8 = 0;
const GATE_IN_FLIGHT: u8 = 1;
const GATE_APPLIED: u8 = 2;

/// The semantic disposition of one tracked change.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Disposition {
    /// Keep inserted or moved-to content and discard deleted or moved-from
    /// content.
    Accept,
    /// Discard inserted or moved-to content and keep deleted or moved-from
    /// content.
    Reject,
}

/// The supported tracked-change families.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChangeKind {
    /// An inline `w:ins` wrapper.
    Insert,
    /// An inline `w:del` wrapper.
    Delete,
    /// A complete `w:moveFrom`/`w:moveTo` pair.
    Move,
}

/// One bounded tracked-change resource policy.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum retained source XML bytes.
    pub max_source_bytes: usize,
    /// Maximum XML reader events.
    pub max_events: usize,
    /// Maximum XML element depth.
    pub max_depth: usize,
    /// Maximum grouped changes retained by one snapshot.
    pub max_changes: usize,
    /// Maximum paired move ranges retained by one snapshot.
    pub max_ranges: usize,
    /// Maximum decoded author, ID, and date bytes retained.
    pub max_metadata_bytes: usize,
    /// Maximum encoded bytes in one XML attribute value.
    pub max_attribute_bytes: usize,
    /// Maximum attributes accepted on one XML element.
    pub max_attributes: usize,
    /// Maximum bytes produced by one committed source rewrite.
    pub max_output_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_source_bytes: 64 * 1024 * 1024,
            max_events: 1_000_000,
            max_depth: 128,
            max_changes: 100_000,
            max_ranges: 100_000,
            max_metadata_bytes: 8 * 1024 * 1024,
            max_attribute_bytes: 64 * 1024,
            max_attributes: 256,
            max_output_bytes: 64 * 1024 * 1024,
        }
    }
}

impl Limits {
    /// Validate finite limits against this owner's hard ceilings.
    ///
    /// # Errors
    ///
    /// Returns an error when a caller requests an unbounded or unsupported
    /// resource policy.
    pub fn validate(self) -> Result<Self> {
        for (resource, value, maximum, nonzero) in [
            (
                "source bytes",
                self.max_source_bytes,
                MAX_SOURCE_BYTES,
                true,
            ),
            ("events", self.max_events, MAX_EVENTS, true),
            ("depth", self.max_depth, MAX_DEPTH, true),
            ("attributes", self.max_attributes, MAX_ATTRIBUTES, true),
            (
                "attribute bytes",
                self.max_attribute_bytes,
                MAX_ATTRIBUTE_BYTES,
                true,
            ),
            ("changes", self.max_changes, MAX_CHANGES, false),
            ("move ranges", self.max_ranges, MAX_RANGES, false),
            (
                "metadata bytes",
                self.max_metadata_bytes,
                MAX_METADATA_BYTES,
                false,
            ),
            (
                "output bytes",
                self.max_output_bytes,
                MAX_OUTPUT_BYTES,
                false,
            ),
        ] {
            if (nonzero && value == 0) || value > maximum {
                return Err(limit_error(resource, value, maximum));
            }
        }
        Ok(self)
    }
}

/// A grouped tracked change in source order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Change {
    kind: ChangeKind,
    id: String,
    author: String,
    date: Option<String>,
    name: Option<String>,
}

impl Change {
    /// Return the change family.
    #[must_use]
    pub const fn kind(&self) -> ChangeKind {
        self.kind
    }

    /// Return the exact decoded `w:id` lexical value from the first owner.
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Return the decoded `w:author` value.
    #[must_use]
    pub fn author(&self) -> &str {
        &self.author
    }

    /// Return the decoded optional `w:date` value.
    #[must_use]
    pub fn date(&self) -> Option<&str> {
        self.date.as_deref()
    }

    /// Return the shared move-container name, when this is a ranged move.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum Side {
    From,
    To,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Span {
    start: usize,
    end: usize,
}

impl Span {
    fn new(start: usize, end: usize) -> Result<Self> {
        if start > end {
            return Err(invalid("tracked revision span starts after its end"));
        }
        Ok(Self { start, end })
    }

    const fn len(self) -> usize {
        self.end - self.start
    }
}

#[derive(Debug, Clone)]
struct NameSpan {
    span: Span,
    replacement: Vec<u8>,
}

#[derive(Debug, Clone)]
struct RawAttribute {
    name: Vec<u8>,
    value: Vec<u8>,
    raw: Vec<u8>,
}

#[derive(Debug, Clone)]
struct NamespaceDeclaration {
    name: Vec<u8>,
    value: Vec<u8>,
    raw: Vec<u8>,
}

#[derive(Debug, Clone)]
struct ScopeAttribute {
    local: Vec<u8>,
    namespace: Vec<u8>,
    value: Vec<u8>,
    raw: Vec<u8>,
}

#[derive(Debug, Clone)]
struct ChildOpening {
    span: Span,
    namespace_declarations: Vec<NamespaceDeclaration>,
    scope_attributes: Vec<ScopeAttribute>,
}

#[derive(Debug, Clone)]
struct Metadata {
    key: String,
    id: String,
    author: String,
    date: Option<String>,
    name: Option<String>,
}

#[derive(Debug, Clone)]
struct Wrapper {
    side: Option<Side>,
    kind: ChangeKind,
    start: Span,
    end: Option<Span>,
    whole: Span,
    deleted_names: Vec<NameSpan>,
    namespace_declarations: Vec<NamespaceDeclaration>,
    scope_attributes: Vec<ScopeAttribute>,
    children: Vec<ChildOpening>,
}

#[derive(Debug, Clone)]
struct MoveRange {
    start_marker: Span,
    end_marker: Span,
    content: Span,
}

#[derive(Debug, Clone)]
struct Group {
    change: Option<Change>,
    kind: Option<ChangeKind>,
    wrappers: Vec<Wrapper>,
    from_range: Option<MoveRange>,
    to_range: Option<MoveRange>,
    order: usize,
}

impl Default for Group {
    fn default() -> Self {
        Self {
            change: None,
            kind: None,
            wrappers: Vec::new(),
            from_range: None,
            to_range: None,
            order: usize::MAX,
        }
    }
}

#[derive(Debug, Clone)]
struct Binding {
    part: String,
    content_type: String,
    topology: Arc<[u8]>,
}

impl PartialEq for Binding {
    fn eq(&self, other: &Self) -> bool {
        self.part == other.part
            && self.content_type == other.content_type
            && self.topology.as_ref() == other.topology.as_ref()
    }
}

impl Eq for Binding {}

impl Binding {
    fn new(part: &PackURI, content_type: &str, topology: Arc<[u8]>) -> Self {
        Self {
            part: part.as_str().to_owned(),
            content_type: content_type.to_owned(),
            topology,
        }
    }
}

#[derive(Debug, Clone)]
struct SnapshotInner {
    source: Arc<Vec<u8>>,
    groups: Arc<Vec<Group>>,
    changes: Arc<Vec<Change>>,
    limits: Limits,
    binding: Option<Binding>,
}

/// An immutable, exact-source tracked-change snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    inner: Arc<SnapshotInner>,
}

impl Snapshot {
    /// Parse one standalone WordprocessingML main-document XML source.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed XML, unsupported property/paragraph
    /// revisions, incomplete move ownership, or exhausted limits.
    pub fn from_xml(source: impl Into<Vec<u8>>) -> Result<Self> {
        Self::from_xml_with_limits(source, Limits::default())
    }

    /// Parse one standalone source under explicit bounded limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the source is outside the supported owner or the
    /// supplied policy.
    pub fn from_xml_with_limits(source: impl Into<Vec<u8>>, limits: Limits) -> Result<Self> {
        let source = Arc::new(source.into());
        Self::from_source(source, limits, None)
    }

    fn from_source(source: Arc<Vec<u8>>, limits: Limits, binding: Option<Binding>) -> Result<Self> {
        let limits = limits.validate()?;
        if source.len() > limits.max_source_bytes {
            return Err(limit_error(
                "source bytes",
                source.len(),
                limits.max_source_bytes,
            ));
        }
        let groups = parse(source.as_slice(), limits)?;
        let mut changes = Vec::new();
        changes
            .try_reserve_exact(groups.len())
            .map_err(|_| Error::RevisionAllocation {
                resource: "tracked revision changes",
            })?;
        for group in &groups {
            changes.push(
                group
                    .change
                    .clone()
                    .ok_or_else(|| invalid("tracked revision group has no public identity"))?,
            );
        }
        Ok(Self {
            inner: Arc::new(SnapshotInner {
                source,
                groups: Arc::new(groups),
                changes: Arc::new(changes),
                limits,
                binding,
            }),
        })
    }

    fn from_blob_with_limits(
        source: Arc<Vec<u8>>,
        limits: Limits,
        binding: Binding,
    ) -> Result<Self> {
        Self::from_source(source, limits, Some(binding))
    }

    /// Borrow the exact source bytes retained by this snapshot.
    #[must_use]
    pub fn source(&self) -> &[u8] {
        self.inner.source.as_slice()
    }

    /// Borrow grouped tracked changes in source order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        self.inner.changes.as_slice()
    }

    /// Return the policy retained by this snapshot.
    #[must_use]
    pub fn limits(&self) -> Limits {
        self.inner.limits
    }

    /// Start a failure-atomic tracked-change disposition transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction::new(self.clone())
    }

    fn source_owner(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.inner.source)
    }

    fn binding(&self) -> Option<&Binding> {
        self.inner.binding.as_ref()
    }

    fn groups(&self) -> &[Group] {
        self.inner.groups.as_slice()
    }
}

/// One failure-atomic tracked-change transaction.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    dispositions: Vec<Option<Disposition>>,
}

impl Transaction {
    fn new(base: Snapshot) -> Self {
        Self {
            dispositions: Vec::new(),
            base,
        }
    }

    /// Borrow the immutable source snapshot.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.base
    }

    /// Whether at least one disposition is staged.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.dispositions.iter().any(Option::is_some)
    }

    /// Stage acceptance of one grouped change.
    ///
    /// # Errors
    ///
    /// Returns an out-of-bounds error for an unknown source-order index.
    pub fn accept(&mut self, index: usize) -> Result<&mut Self> {
        self.set(index, Disposition::Accept)
    }

    /// Stage rejection of one grouped change.
    ///
    /// # Errors
    ///
    /// Returns an out-of-bounds error for an unknown source-order index.
    pub fn reject(&mut self, index: usize) -> Result<&mut Self> {
        self.set(index, Disposition::Reject)
    }

    /// Stage one explicit disposition.
    ///
    /// # Errors
    ///
    /// Returns an out-of-bounds error for an unknown source-order index.
    pub fn set(&mut self, index: usize, disposition: Disposition) -> Result<&mut Self> {
        let len = self.base.changes().len();
        if index >= len {
            return Err(Error::OutOfBounds {
                object: "tracked revision",
                index,
                len,
            });
        }
        self.ensure_dispositions(index + 1)?;
        let slot = &mut self.dispositions[index];
        *slot = Some(disposition);
        Ok(self)
    }

    /// Stage acceptance of every grouped change.
    ///
    /// # Errors
    ///
    /// Returns an allocation error only if the internal disposition table could
    /// not be retained; the table is pre-sized by the source snapshot.
    pub fn accept_all(&mut self) -> Result<&mut Self> {
        self.ensure_dispositions(self.base.changes().len())?;
        self.dispositions.fill(Some(Disposition::Accept));
        Ok(self)
    }

    /// Stage rejection of every grouped change.
    ///
    /// # Errors
    ///
    /// Returns an allocation error only if the internal disposition table could
    /// not be retained; the table is pre-sized by the source snapshot.
    pub fn reject_all(&mut self) -> Result<&mut Self> {
        self.ensure_dispositions(self.base.changes().len())?;
        self.dispositions.fill(Some(Disposition::Reject));
        Ok(self)
    }

    /// Validate and materialize a fully reparsed candidate and reversible patch.
    ///
    /// The source snapshot is never mutated. A failed commit leaves this
    /// transaction available for retry.
    ///
    /// # Errors
    ///
    /// Returns an XML, ownership, or resource-limit error.
    pub fn commit(&self) -> Result<Commit> {
        let before = self.base.source_owner();
        let source_limits = self.base.limits();
        if !self.is_changed() {
            return Ok(Commit::new(
                self.base.clone(),
                Arc::clone(&before),
                before,
                source_limits,
                source_limits,
            ));
        }

        let mut splices = Vec::new();
        splices
            .try_reserve(self.dispositions.iter().flatten().count().saturating_mul(8))
            .map_err(|_| Error::RevisionAllocation {
                resource: "tracked revision source splices",
            })?;
        for index in 0..self.base.groups().len() {
            let Some(disposition) = self.dispositions.get(index).and_then(Option::as_ref) else {
                continue;
            };
            let group = self
                .base
                .groups()
                .get(index)
                .ok_or_else(|| invalid("tracked revision transaction index is stale"))?;
            plan_group(self.base.source(), group, *disposition, &mut splices)?;
        }
        let after = apply_splices(
            self.base.source(),
            &mut splices,
            source_limits.max_output_bytes,
        )?;
        let after = Arc::new(after);
        let mut target_limits = source_limits;
        target_limits.max_source_bytes = target_limits.max_source_bytes.max(after.len());
        let target_limits = target_limits.validate()?;
        let snapshot = Snapshot::from_source(
            Arc::clone(&after),
            target_limits,
            self.base.binding().cloned(),
        )?;
        Ok(Commit::new(
            snapshot,
            before,
            after,
            source_limits,
            target_limits,
        ))
    }

    fn ensure_dispositions(&mut self, len: usize) -> Result<()> {
        if self.dispositions.len() >= len {
            return Ok(());
        }
        self.dispositions
            .try_reserve_exact(len - self.dispositions.len())
            .map_err(|_| Error::RevisionAllocation {
                resource: "tracked revision dispositions",
            })?;
        self.dispositions.resize(len, None);
        Ok(())
    }
}

/// A successful tracked-change disposition commit.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    fn new(
        snapshot: Snapshot,
        before: Arc<Vec<u8>>,
        after: Arc<Vec<u8>>,
        source_limits: Limits,
        target_limits: Limits,
    ) -> Self {
        let binding = snapshot.binding().cloned();
        Self {
            snapshot,
            patch: Patch::new(before, after, source_limits, target_limits, binding),
        }
    }

    /// Whether the disposition changes any source byte.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Borrow the fully reparsed candidate.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the candidate snapshot out of the commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Move the patch out of the commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// An exact, reversible, source-bound XML replacement.
#[derive(Debug, Clone)]
pub struct Patch {
    source_fingerprint: u64,
    target_fingerprint: u64,
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
    source_limits: Limits,
    target_limits: Limits,
    binding: Option<Binding>,
    gate: Arc<AtomicU8>,
}

impl Patch {
    fn new(
        before: Arc<Vec<u8>>,
        after: Arc<Vec<u8>>,
        source_limits: Limits,
        target_limits: Limits,
        binding: Option<Binding>,
    ) -> Self {
        Self {
            source_fingerprint: fingerprint(before.as_slice()),
            target_fingerprint: fingerprint(after.as_slice()),
            before,
            after,
            source_limits,
            target_limits,
            binding,
            gate: Arc::new(AtomicU8::new(GATE_READY)),
        }
    }

    /// Return the compact source fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the compact target fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Borrow the exact source bytes required by this patch.
    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        self.before.as_slice()
    }

    /// Borrow the exact target bytes produced by this patch.
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        self.after.as_slice()
    }

    /// Return the target parsing/publication policy.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.target_limits
    }

    /// Whether this patch preserves every source byte.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_slice() == self.after.as_slice()
    }

    /// Whether this patch or one of its clones has been successfully applied.
    #[must_use]
    pub fn is_applied(&self) -> bool {
        self.gate.load(Ordering::Acquire) == GATE_APPLIED
    }

    /// Apply this patch to an exact detached or package-bound snapshot.
    ///
    /// # Errors
    ///
    /// Returns an exact-source, policy, binding, or target-parse error.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        let claim = self.claim_publication()?;
        if source.inner.limits != self.source_limits
            || source.binding() != self.binding.as_ref()
            || source.source() != self.before.as_slice()
            || fingerprint(source.source()) != self.source_fingerprint
        {
            return Err(invalid(
                "tracked revision patch source does not match its exact precondition",
            ));
        }
        let candidate = Snapshot::from_source(
            Arc::clone(&self.after),
            self.target_limits,
            self.binding.clone(),
        )?;
        claim.finalize();
        Ok(candidate)
    }

    /// Build a fresh inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self::new(
            Arc::clone(&self.after),
            Arc::clone(&self.before),
            self.target_limits,
            self.source_limits,
            self.binding.clone(),
        )
    }

    fn claim_publication(&self) -> Result<PublicationClaim> {
        match self.gate.compare_exchange(
            GATE_READY,
            GATE_IN_FLIGHT,
            Ordering::AcqRel,
            Ordering::Acquire,
        ) {
            Ok(_) => Ok(PublicationClaim {
                gate: Arc::clone(&self.gate),
                finalized: false,
            }),
            Err(GATE_IN_FLIGHT) => Err(invalid("tracked revision patch publication is in flight")),
            Err(_) => Err(invalid("tracked revision patch has already been applied")),
        }
    }
}

struct PublicationClaim {
    gate: Arc<AtomicU8>,
    finalized: bool,
}

impl PublicationClaim {
    fn finalize(mut self) {
        self.gate.store(GATE_APPLIED, Ordering::Release);
        self.finalized = true;
    }
}

impl Drop for PublicationClaim {
    fn drop(&mut self) {
        if !self.finalized {
            self.gate.store(GATE_READY, Ordering::Release);
        }
    }
}

impl Package {
    /// Read tracked inline and move revisions from the resolved main document.
    ///
    /// # Errors
    ///
    /// Returns an error when the main document is stale, unsupported, malformed,
    /// or outside the requested policy.
    pub fn tracked_revisions(&self) -> Result<Snapshot> {
        self.tracked_revisions_with_limits(Limits::default())
    }

    /// Read tracked revisions from the main document under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the main document is stale, unsupported, malformed,
    /// or outside the requested policy.
    pub fn tracked_revisions_with_limits(&self, limits: Limits) -> Result<Snapshot> {
        let limits = limits.validate()?;
        self.ensure_conflict_opc_current("tracked_revisions_with_limits")?;
        let story_limits = package_story_limits(limits);
        let inventory = self.story_inventory_with_limits(story_limits)?;
        let story = inventory
            .get(inventory.main())
            .ok_or_else(|| invalid("resolved main document is missing from the story inventory"))?;
        let binding = Binding::new(
            story.part(),
            story.content_type(),
            inventory.topology().shared(),
        );
        Snapshot::from_blob_with_limits(story.source_arc(), limits, binding)
    }

    /// Publish a prepared tracked-revision disposition commit.
    ///
    /// # Errors
    ///
    /// Returns an exact-source, signature-policy, stale-package, or target
    /// validation error without publishing a partial package mutation.
    pub fn apply_tracked_revisions(&mut self, commit: &Commit) -> Result<Snapshot> {
        self.apply_tracked_revision_patch(commit.patch())
    }

    /// Publish a reversible tracked-revision patch against the resolved main
    /// document owner.
    ///
    /// Exact no-ops avoid cloning the OPC package. Changed publication checks
    /// the source and target before entering the package clone-and-publish seam;
    /// the seam itself repeats the source check before replacing one main-part
    /// blob and preserving every unrelated relationship and part.
    ///
    /// # Errors
    ///
    /// Returns an exact-source, signature-policy, stale-package, or target
    /// validation error without publishing a partial package mutation.
    pub fn apply_tracked_revision_patch(&mut self, patch: &Patch) -> Result<Snapshot> {
        let claim = patch.claim_publication()?;
        self.ensure_conflict_opc_current("apply_tracked_revision_patch")?;
        let binding = patch
            .binding
            .as_ref()
            .ok_or_else(|| invalid("detached tracked revision patch cannot target a package"))?
            .clone();
        let source_limits = patch.source_limits;
        let target_limits = patch.target_limits;
        let story_limits = package_story_limits(source_limits);
        let inventory = self.story_inventory_with_limits(story_limits)?;
        if inventory.topology().as_bytes() != binding.topology.as_ref() {
            return Err(invalid("tracked revision package owner topology is stale"));
        }
        let target = PackURI::new(binding.part.clone())
            .map_err(|error| Error::InvalidUri(format!("tracked revision owner URI: {error}")))?;
        let story = inventory
            .get(&target)
            .ok_or_else(|| invalid("tracked revision patch source owner is stale"))?;
        if story.content_type() != binding.content_type || story.source() != patch.before.as_slice()
        {
            return Err(invalid("tracked revision patch source owner is stale"));
        }
        let current = Snapshot::from_blob_with_limits(
            story.source_arc(),
            source_limits,
            Binding {
                part: binding.part.clone(),
                content_type: binding.content_type.clone(),
                topology: Arc::clone(&binding.topology),
            },
        )?;
        if patch.is_noop() {
            claim.finalize();
            return Ok(current);
        }
        let expected = Snapshot::from_source(
            Arc::clone(&patch.after),
            target_limits,
            Some(binding.clone()),
        )?;
        let before = Arc::clone(&patch.before);
        let replacement = Arc::clone(&patch.after);
        let expected_topology = Arc::clone(&binding.topology);
        let published =
            self.edit_semantic_opc("apply_tracked_revision_patch", move |candidate| {
                let staged = capture(candidate, story_limits)?;
                if staged.topology().as_bytes() != expected_topology.as_ref() {
                    return Err(invalid("tracked revision package owner topology is stale"));
                }
                let part = staged
                    .get(&target)
                    .ok_or_else(|| invalid("tracked revision patch source owner is stale"))?;
                if part.content_type() != binding.content_type || part.source() != before.as_slice()
                {
                    return Err(invalid("tracked revision patch source owner is stale"));
                }
                candidate
                    .get_part_mut(&target)?
                    .set_blob_shared(replacement);
                Ok(expected)
            })?;
        claim.finalize();
        Ok(published)
    }
}

#[derive(Debug, Clone)]
struct RawWrapper {
    side: Option<Side>,
    kind: ChangeKind,
    metadata: Metadata,
    start: Span,
    end: Option<Span>,
    whole: Span,
    deleted_names: Vec<NameSpan>,
    namespace_declarations: Vec<NamespaceDeclaration>,
    scope_attributes: Vec<ScopeAttribute>,
    children: Vec<ChildOpening>,
}

#[derive(Debug, Clone)]
struct RawRange {
    side: Side,
    metadata: Metadata,
    name: String,
    start_marker: Span,
    end_marker: Span,
    content: Span,
}

#[derive(Debug, Clone)]
struct Frame {
    id: usize,
    namespace: NamespaceKind,
    local: Vec<u8>,
    wrapper: Option<usize>,
    mce_scope: bool,
    run_rpr_seen: bool,
    run_content_started: bool,
    run_properties_seen: u64,
    run_properties_change_seen: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Word,
    MarkupCompatibility,
    Other,
    Unbound,
}

impl NamespaceKind {
    const fn is_word(self) -> bool {
        matches!(self, Self::Word)
    }
}

#[derive(Debug, Clone)]
struct OpenRange {
    side: Side,
    metadata: Metadata,
    name: String,
    marker: Span,
    order: usize,
    parent_frame: Option<usize>,
}

fn parse(source: &[u8], limits: Limits) -> Result<Vec<Group>> {
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut frames = Vec::<Frame>::new();
    frames
        .try_reserve(limits.max_depth.min(256))
        .map_err(|_| Error::RevisionAllocation {
            resource: "tracked revision XML stack",
        })?;
    let mut wrappers = Vec::<RawWrapper>::new();
    let mut ranges = Vec::<RawRange>::new();
    let mut open_ranges = HashMap::<(Side, String), OpenRange>::new();
    let mut range_start_ids = HashSet::<String>::new();
    let mut events = 0usize;
    let mut metadata_bytes = 0usize;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut mce_depth = 0usize;
    let mut order = 0usize;
    let mut frame_id = 0usize;
    let mut saw_custom_xml_range = false;

    loop {
        let begin = position(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let end = position(&reader)?;
        if !matches!(event, Event::Eof) {
            events = events
                .checked_add(1)
                .ok_or_else(|| invalid("tracked revision event counter overflow"))?;
            if events > limits.max_events {
                return Err(limit_error("events", events, limits.max_events));
            }
        }
        let resolver = reader.resolver().clone();
        let (resolved, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                reject_unknown_namespace(&resolved)?;
                if root_closed {
                    return Err(invalid("tracked revision XML has multiple root elements"));
                }
                if depth == 0 {
                    if root_seen {
                        return Err(invalid("tracked revision XML has multiple root elements"));
                    }
                    root_seen = true;
                    if !is_word_namespace(&resolved) || element.local_name().as_ref() != b"document"
                    {
                        return Err(invalid(
                            "tracked revision owner requires a WordprocessingML document root",
                        ));
                    }
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("tracked revision XML depth overflow"))?;
                if depth > limits.max_depth {
                    return Err(limit_error("depth", depth, limits.max_depth));
                }
                validate_attributes(
                    &element,
                    &resolver,
                    reader.decoder(),
                    limits,
                    &mut metadata_bytes,
                    false,
                )?;
                let namespace = namespace_kind(&resolved);
                let local = element.local_name().as_ref().to_vec();
                if is_custom_xml_range_marker(local.as_slice(), namespace) {
                    saw_custom_xml_range = true;
                }
                validate_run_track_change_child(frames.last(), local.as_slice(), namespace)?;
                validate_run_child(
                    &mut frames,
                    !open_ranges.is_empty(),
                    local.as_slice(),
                    namespace,
                )?;
                validate_run_properties_child(
                    &mut frames,
                    !open_ranges.is_empty(),
                    local.as_slice(),
                    namespace,
                )?;
                reject_active_markup_compatibility(namespace, &frames, &open_ranges)?;
                let parent = frames.last();
                let parent_wrapper = parent.and_then(|frame| frame.wrapper);
                if classify_range(local.as_slice(), namespace).is_some()
                    && !legal_run_level_parent(parent)
                {
                    return Err(invalid(
                        "tracked move range marker has an unsupported paragraph or property owner",
                    ));
                }
                let current_mce = mce_depth != 0;
                let mce_scope = namespace == NamespaceKind::MarkupCompatibility
                    && matches!(
                        local.as_slice(),
                        b"AlternateContent" | b"Choice" | b"Fallback"
                    );
                let active_mce = current_mce || mce_scope;
                if (frames.iter().any(|frame| frame.wrapper.is_some()) || !open_ranges.is_empty())
                    && namespace.is_word()
                    && unsupported_dependency(local.as_slice())
                {
                    return Err(invalid(
                        "tracked revision content has an unsupported dependency-bearing owner",
                    ));
                }
                if namespace.is_word() && unsupported_property_revision(local.as_slice()) {
                    return Err(invalid(
                        "tracked revision property changes are outside the authoring owner",
                    ));
                }
                let classified = classify_wrapper(local.as_slice(), namespace, parent)?;
                if classify_range(local.as_slice(), namespace).is_some() {
                    return Err(invalid(
                        "tracked move range markers must be empty WordprocessingML elements",
                    ));
                }
                let wrapper = if let Some((kind, side)) = classified {
                    if active_mce {
                        return Err(invalid(
                            "tracked revisions inside markup-compatibility alternatives are unsupported",
                        ));
                    }
                    if frames.iter().any(|frame| frame.wrapper.is_some()) {
                        return Err(invalid(
                            "nested WordprocessingML tracked changes are unsupported",
                        ));
                    }
                    if has_field_or_control_ancestor(&frames) {
                        return Err(invalid(
                            "tracked revisions inside fields or structured document controls are unsupported",
                        ));
                    }
                    let metadata = parse_metadata(
                        source,
                        begin,
                        end,
                        &element,
                        &resolver,
                        reader.decoder(),
                        limits,
                        &mut metadata_bytes,
                        false,
                    )?;
                    let (namespace_declarations, scope_attributes) = collect_hoisted_attributes(
                        source,
                        begin,
                        end,
                        &element,
                        &resolver,
                        reader.decoder(),
                        limits,
                    )?;
                    let index = wrappers.len();
                    push_bounded(
                        &mut wrappers,
                        RawWrapper {
                            side,
                            kind,
                            metadata,
                            start: Span::new(begin, end)?,
                            end: None,
                            whole: Span::new(begin, end)?,
                            deleted_names: Vec::new(),
                            namespace_declarations,
                            scope_attributes,
                            children: Vec::new(),
                        },
                        "tracked revision wrappers",
                    )?;
                    Some(index)
                } else {
                    None
                };
                if let Some(parent_wrapper) = parent_wrapper {
                    let (namespace_declarations, scope_attributes) = collect_hoisted_attributes(
                        source,
                        begin,
                        end,
                        &element,
                        &resolver,
                        reader.decoder(),
                        limits,
                    )?;
                    let child = ChildOpening {
                        span: Span::new(begin, end)?,
                        namespace_declarations,
                        scope_attributes,
                    };
                    let wrapper = wrappers
                        .get_mut(parent_wrapper)
                        .ok_or_else(|| invalid("tracked revision parent wrapper is stale"))?;
                    push_bounded(
                        &mut wrapper.children,
                        child,
                        "tracked revision wrapper child openings",
                    )?;
                }
                if local.as_slice() == b"delText" && namespace.is_word() {
                    let wrapper = frames
                        .iter()
                        .rev()
                        .find_map(|frame| frame.wrapper)
                        .ok_or_else(|| {
                            invalid("w:delText is outside a supported deleted revision")
                        })?;
                    if !matches!(wrappers[wrapper].kind, ChangeKind::Delete)
                        && wrappers[wrapper].side != Some(Side::From)
                    {
                        return Err(invalid("w:delText is outside a deleted revision side"));
                    }
                    let name = name_span(source, begin, end, false)?;
                    push_bounded(
                        &mut wrappers[wrapper].deleted_names,
                        name,
                        "tracked revision deleted text names",
                    )?;
                }
                let frame = Frame {
                    id: frame_id,
                    namespace,
                    local,
                    wrapper,
                    mce_scope,
                    run_rpr_seen: false,
                    run_content_started: false,
                    run_properties_seen: 0,
                    run_properties_change_seen: false,
                };
                frame_id = frame_id
                    .checked_add(1)
                    .ok_or_else(|| invalid("tracked revision frame counter overflow"))?;
                if mce_scope {
                    mce_depth = mce_depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("markup-compatibility depth overflow"))?;
                }
                frames.push(frame);
            },
            Event::Empty(element) => {
                reject_unknown_namespace(&resolved)?;
                if depth == 0 {
                    if root_seen {
                        return Err(invalid("tracked revision XML has multiple root elements"));
                    }
                    root_seen = true;
                    root_closed = true;
                    if !is_word_namespace(&resolved) || element.local_name().as_ref() != b"document"
                    {
                        return Err(invalid(
                            "tracked revision owner requires a WordprocessingML document root",
                        ));
                    }
                }
                validate_attributes(
                    &element,
                    &resolver,
                    reader.decoder(),
                    limits,
                    &mut metadata_bytes,
                    false,
                )?;
                let namespace = namespace_kind(&resolved);
                let local = element.local_name().as_ref().to_vec();
                if is_custom_xml_range_marker(local.as_slice(), namespace) {
                    saw_custom_xml_range = true;
                }
                validate_run_track_change_child(frames.last(), local.as_slice(), namespace)?;
                validate_run_child(
                    &mut frames,
                    !open_ranges.is_empty(),
                    local.as_slice(),
                    namespace,
                )?;
                validate_run_properties_child(
                    &mut frames,
                    !open_ranges.is_empty(),
                    local.as_slice(),
                    namespace,
                )?;
                reject_active_markup_compatibility(namespace, &frames, &open_ranges)?;
                let parent = frames.last();
                let parent_wrapper = parent.and_then(|frame| frame.wrapper);
                if classify_range(local.as_slice(), namespace).is_some()
                    && !legal_run_level_parent(parent)
                {
                    return Err(invalid(
                        "tracked move range marker has an unsupported paragraph or property owner",
                    ));
                }
                let active_mce = mce_depth != 0;
                if (frames.iter().any(|frame| frame.wrapper.is_some()) || !open_ranges.is_empty())
                    && namespace.is_word()
                    && unsupported_dependency(local.as_slice())
                {
                    return Err(invalid(
                        "tracked revision content has an unsupported dependency-bearing owner",
                    ));
                }
                if namespace.is_word() && unsupported_property_revision(local.as_slice()) {
                    return Err(invalid(
                        "tracked revision property changes are outside the authoring owner",
                    ));
                }
                if let Some((kind, side)) = classify_wrapper(local.as_slice(), namespace, parent)? {
                    if active_mce {
                        return Err(invalid(
                            "tracked revisions inside markup-compatibility alternatives are unsupported",
                        ));
                    }
                    if frames.iter().any(|frame| frame.wrapper.is_some()) {
                        return Err(invalid(
                            "nested WordprocessingML tracked changes are unsupported",
                        ));
                    }
                    if has_field_or_control_ancestor(&frames) {
                        return Err(invalid(
                            "tracked revisions inside fields or structured document controls are unsupported",
                        ));
                    }
                    let metadata = parse_metadata(
                        source,
                        begin,
                        end,
                        &element,
                        &resolver,
                        reader.decoder(),
                        limits,
                        &mut metadata_bytes,
                        false,
                    )?;
                    let (namespace_declarations, scope_attributes) = collect_hoisted_attributes(
                        source,
                        begin,
                        end,
                        &element,
                        &resolver,
                        reader.decoder(),
                        limits,
                    )?;
                    push_bounded(
                        &mut wrappers,
                        RawWrapper {
                            side,
                            kind,
                            metadata,
                            start: Span::new(begin, end)?,
                            end: None,
                            whole: Span::new(begin, end)?,
                            deleted_names: Vec::new(),
                            namespace_declarations,
                            scope_attributes,
                            children: Vec::new(),
                        },
                        "tracked revision wrappers",
                    )?;
                } else if let Some((side, is_start)) = classify_range(local.as_slice(), namespace) {
                    if active_mce {
                        return Err(invalid(
                            "tracked move ranges inside markup-compatibility alternatives are unsupported",
                        ));
                    }
                    if frames.iter().any(|frame| frame.wrapper.is_some()) {
                        return Err(invalid(
                            "tracked move range markers inside a tracked wrapper are unsupported",
                        ));
                    }
                    if has_field_or_control_ancestor(&frames) {
                        return Err(invalid(
                            "tracked move ranges inside fields or structured document controls are unsupported",
                        ));
                    }
                    let metadata = if is_start {
                        parse_metadata(
                            source,
                            begin,
                            end,
                            &element,
                            &resolver,
                            reader.decoder(),
                            limits,
                            &mut metadata_bytes,
                            true,
                        )?
                    } else {
                        parse_range_end_metadata(
                            &element,
                            &resolver,
                            reader.decoder(),
                            limits,
                            &mut metadata_bytes,
                        )?
                    };
                    let key = (side, metadata.key.clone());
                    if is_start {
                        if range_start_ids.contains(&metadata.key) {
                            return Err(invalid("duplicate tracked move range start ID"));
                        }
                        if open_ranges.contains_key(&key) {
                            return Err(invalid("duplicate tracked move range start"));
                        }
                        if !open_ranges.is_empty() {
                            return Err(invalid("tracked move ranges must not nest or cross"));
                        }
                        if ranges.len() >= limits.max_ranges {
                            return Err(limit_error(
                                "move ranges",
                                ranges.len() + 1,
                                limits.max_ranges,
                            ));
                        }
                        range_start_ids
                            .try_reserve(1)
                            .map_err(|_| Error::RevisionAllocation {
                                resource: "tracked revision move range IDs",
                            })?;
                        range_start_ids.insert(metadata.key.clone());
                        open_ranges
                            .try_reserve(1)
                            .map_err(|_| Error::RevisionAllocation {
                                resource: "tracked revision move range owners",
                            })?;
                        let name = metadata
                            .name
                            .clone()
                            .ok_or_else(|| invalid("tracked move range has no name"))?;
                        open_ranges.insert(
                            key,
                            OpenRange {
                                side,
                                metadata,
                                name,
                                marker: Span::new(begin, end)?,
                                order,
                                parent_frame: frames.last().map(|frame| frame.id),
                            },
                        );
                    } else {
                        let open = open_ranges
                            .remove(&key)
                            .ok_or_else(|| invalid("orphaned tracked move range end"))?;
                        if open.parent_frame != frames.last().map(|frame| frame.id) {
                            return Err(invalid("tracked move range crosses XML ownership"));
                        }
                        if open_ranges.values().any(|other| other.order > open.order) {
                            return Err(invalid("tracked move ranges overlap or cross"));
                        }
                        if metadata.key != open.metadata.key {
                            return Err(invalid("tracked move range IDs do not match"));
                        }
                        push_bounded(
                            &mut ranges,
                            RawRange {
                                side: open.side,
                                metadata: open.metadata,
                                name: open.name,
                                start_marker: open.marker,
                                end_marker: Span::new(begin, end)?,
                                content: Span::new(open.marker.end, begin)?,
                            },
                            "tracked revision move ranges",
                        )?;
                    }
                } else if namespace.is_word() && local.as_slice() == b"delText" {
                    let wrapper = frames
                        .iter()
                        .rev()
                        .find_map(|frame| frame.wrapper)
                        .ok_or_else(|| {
                            invalid("w:delText is outside a supported deleted revision")
                        })?;
                    if !matches!(wrappers[wrapper].kind, ChangeKind::Delete)
                        && wrappers[wrapper].side != Some(Side::From)
                    {
                        return Err(invalid("w:delText is outside a deleted revision side"));
                    }
                    let name = name_span(source, begin, end, false)?;
                    push_bounded(
                        &mut wrappers[wrapper].deleted_names,
                        name,
                        "tracked revision deleted text names",
                    )?;
                }
                if let Some(parent_wrapper) = parent_wrapper {
                    let (namespace_declarations, scope_attributes) = collect_hoisted_attributes(
                        source,
                        begin,
                        end,
                        &element,
                        &resolver,
                        reader.decoder(),
                        limits,
                    )?;
                    let child = ChildOpening {
                        span: Span::new(begin, end)?,
                        namespace_declarations,
                        scope_attributes,
                    };
                    let wrapper = wrappers
                        .get_mut(parent_wrapper)
                        .ok_or_else(|| invalid("tracked revision parent wrapper is stale"))?;
                    push_bounded(
                        &mut wrapper.children,
                        child,
                        "tracked revision wrapper child openings",
                    )?;
                }
                order = order
                    .checked_add(1)
                    .ok_or_else(|| invalid("tracked revision source order overflow"))?;
            },
            Event::End(element) => {
                reject_unknown_namespace(&resolved)?;
                if classify_range(element.local_name().as_ref(), namespace_kind(&resolved))
                    .is_some()
                {
                    return Err(invalid(
                        "tracked move range markers must be empty WordprocessingML elements",
                    ));
                }
                let frame = frames
                    .pop()
                    .ok_or_else(|| invalid("tracked revision XML has an unmatched end element"))?;
                if frame.local.as_slice() != element.local_name().as_ref()
                    || frame.namespace != namespace_kind(&resolved)
                {
                    return Err(invalid(
                        "tracked revision XML element nesting is inconsistent",
                    ));
                }
                if let Some(index) = frame.wrapper {
                    let end_span = Span::new(begin, end)?;
                    let wrapper = wrappers
                        .get_mut(index)
                        .ok_or_else(|| invalid("tracked revision wrapper index is stale"))?;
                    wrapper.end = Some(end_span);
                    wrapper.whole = Span::new(wrapper.start.start, end)?;
                }
                if frame.namespace.is_word() && frame.local.as_slice() == b"delText" {
                    let wrapper = frames
                        .iter()
                        .rev()
                        .find_map(|frame| frame.wrapper)
                        .ok_or_else(|| {
                            invalid("w:delText is outside a supported deleted revision")
                        })?;
                    if !matches!(wrappers[wrapper].kind, ChangeKind::Delete)
                        && wrappers[wrapper].side != Some(Side::From)
                    {
                        return Err(invalid("w:delText is outside a deleted revision side"));
                    }
                    let name = name_span(source, begin, end, true)?;
                    push_bounded(
                        &mut wrappers[wrapper].deleted_names,
                        name,
                        "tracked revision deleted text names",
                    )?;
                }
                if frame.mce_scope {
                    mce_depth = mce_depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("markup-compatibility depth underflow"))?;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("tracked revision XML depth underflow"))?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                if depth == 0 && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("tracked revision XML has top-level text"));
                }
                if direct_run_track_change_frame(&frames)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid(
                        "tracked revision wrapper has text outside a run-content element",
                    ));
                }
                reject_run_parent_text(&frames, !open_ranges.is_empty(), text.as_ref())?;
            },
            Event::CData(text) => {
                if depth == 0 && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("tracked revision XML has top-level text"));
                }
                if direct_run_track_change_frame(&frames)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid(
                        "tracked revision wrapper has text outside a run-content element",
                    ));
                }
                reject_run_parent_text(&frames, !open_ranges.is_empty(), text.as_ref())?;
            },
            Event::Eof => {
                if !frames.is_empty() || depth != 0 {
                    return Err(invalid("tracked revision XML has unterminated elements"));
                }
                if !root_seen || !root_closed {
                    return Err(invalid(
                        "tracked revision XML has no complete document root",
                    ));
                }
                if !open_ranges.is_empty() {
                    return Err(invalid("tracked revision move range has no end"));
                }
                break;
            },
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(invalid(
                    "DTD and general entity references are unsupported in a revision owner",
                ));
            },
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
        }
    }
    if saw_custom_xml_range && (!wrappers.is_empty() || !ranges.is_empty()) {
        return Err(invalid(
            "tracked revisions with custom XML range closures are outside the authoring owner",
        ));
    }
    build_groups(wrappers, ranges, limits)
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum GroupKey {
    Ordinary(ChangeKind, String),
    MoveRange(String),
}

fn build_groups(
    wrappers: Vec<RawWrapper>,
    ranges: Vec<RawRange>,
    limits: Limits,
) -> Result<Vec<Group>> {
    let mut groups = HashMap::<GroupKey, Group>::new();

    // Ranges are owned by their shared w:name. Their marker IDs are only the
    // local start/end closure token and must never pair source and destination
    // containers.
    for range in ranges {
        let key = GroupKey::MoveRange(range.name.clone());
        ensure_group_slot(&mut groups, &key, limits.max_changes)?;
        let group = groups.entry(key).or_default();
        group.kind = Some(ChangeKind::Move);
        merge_range_metadata(group, &range.metadata, &range.name)?;
        group.order = group.order.min(range.start_marker.start);
        let target = match range.side {
            Side::From => &mut group.from_range,
            Side::To => &mut group.to_range,
        };
        let move_range = MoveRange {
            start_marker: range.start_marker,
            end_marker: range.end_marker,
            content: range.content,
        };
        if target.replace(move_range).is_some() {
            return Err(invalid("duplicate tracked move range owner"));
        }
    }

    for wrapper in wrappers {
        let metadata = wrapper.metadata.clone();
        let key = if wrapper.kind == ChangeKind::Move {
            let name = containing_move_range(&groups, &wrapper)?.ok_or_else(|| {
                invalid("tracked move wrapper must be contained by its complete named move range")
            })?;
            GroupKey::MoveRange(name)
        } else {
            GroupKey::Ordinary(wrapper.kind, metadata.key.clone())
        };
        ensure_group_slot(&mut groups, &key, limits.max_changes)?;
        let group = groups.entry(key).or_default();
        group.kind = Some(wrapper.kind);
        if wrapper.kind == ChangeKind::Move {
            merge_move_wrapper_metadata(group, &metadata)?;
        } else {
            merge_plain_metadata(group, &metadata)?;
        }
        append_wrapper(group, wrapper)?;
    }

    let mut output = Vec::new();
    output
        .try_reserve(groups.len())
        .map_err(|_| Error::RevisionAllocation {
            resource: "tracked revision groups",
        })?;
    for (_, mut group) in groups {
        let kind = group
            .kind
            .ok_or_else(|| invalid("tracked revision group has no kind"))?;
        let change = group
            .change
            .as_mut()
            .ok_or_else(|| invalid("tracked revision group has no metadata"))?;
        change.kind = kind;
        if kind == ChangeKind::Move {
            let has_from_wrapper = group
                .wrappers
                .iter()
                .any(|wrapper| wrapper.side == Some(Side::From));
            let has_to_wrapper = group
                .wrappers
                .iter()
                .any(|wrapper| wrapper.side == Some(Side::To));
            let has_from_range = group.from_range.is_some();
            let has_to_range = group.to_range.is_some();
            if has_from_range != has_to_range {
                return Err(invalid(
                    "tracked move range owner is missing its opposite range",
                ));
            }
            if !has_from_wrapper || !has_to_wrapper {
                return Err(invalid("tracked move owner is missing its opposite side"));
            }
        } else if group.wrappers.iter().any(|wrapper| wrapper.side.is_some()) {
            return Err(invalid("ordinary tracked revision has move-side metadata"));
        }
        push_bounded(&mut output, group, "tracked revision groups")?;
    }
    output.sort_unstable_by_key(|group| group.order);
    Ok(output)
}

fn ensure_group_slot(
    groups: &mut HashMap<GroupKey, Group>,
    key: &GroupKey,
    limit: usize,
) -> Result<()> {
    if groups.contains_key(key) {
        return Ok(());
    }
    let next = groups
        .len()
        .checked_add(1)
        .ok_or_else(|| limit_error("changes", usize::MAX, limit))?;
    if next > limit {
        return Err(limit_error("changes", next, limit));
    }
    groups
        .try_reserve(1)
        .map_err(|_| Error::RevisionAllocation {
            resource: "tracked revision groups",
        })?;
    Ok(())
}

fn containing_move_range(
    groups: &HashMap<GroupKey, Group>,
    wrapper: &RawWrapper,
) -> Result<Option<String>> {
    let mut found = None;
    for (key, group) in groups {
        let GroupKey::MoveRange(name) = key else {
            continue;
        };
        let range = match wrapper.side {
            Some(Side::From) => group.from_range.as_ref(),
            Some(Side::To) => group.to_range.as_ref(),
            None => None,
        };
        if range.is_some_and(|range| {
            wrapper.whole.start >= range.content.start && wrapper.whole.end <= range.content.end
        }) {
            if found.is_some() {
                return Err(invalid("tracked move wrapper belongs to multiple ranges"));
            }
            found = Some(name.clone());
        }
    }
    Ok(found)
}

fn append_wrapper(group: &mut Group, wrapper: RawWrapper) -> Result<()> {
    group.order = group.order.min(wrapper.start.start);
    push_bounded(
        &mut group.wrappers,
        Wrapper {
            side: wrapper.side,
            kind: wrapper.kind,
            start: wrapper.start,
            end: wrapper.end,
            whole: wrapper.whole,
            deleted_names: wrapper.deleted_names,
            namespace_declarations: wrapper.namespace_declarations,
            scope_attributes: wrapper.scope_attributes,
            children: wrapper.children,
        },
        "tracked revision group wrappers",
    )
}

fn merge_plain_metadata(group: &mut Group, metadata: &Metadata) -> Result<()> {
    if let Some(change) = &group.change {
        if normalize_integer(&change.id)? != metadata.key
            || change.author != metadata.author
            || change.date.as_deref() != metadata.date.as_deref()
        {
            return Err(invalid("tracked revision owner metadata is inconsistent"));
        }
    } else {
        group.change = Some(Change {
            kind: ChangeKind::Insert,
            id: metadata.id.clone(),
            author: metadata.author.clone(),
            date: metadata.date.clone(),
            name: None,
        });
    }
    Ok(())
}

fn merge_range_metadata(group: &mut Group, metadata: &Metadata, name: &str) -> Result<()> {
    if metadata.name.as_deref() != Some(name) {
        return Err(invalid("tracked move range name is inconsistent"));
    }
    if let Some(change) = &group.change {
        if change.name.as_deref() != Some(name)
            || change.author != metadata.author
            || change.date.as_deref() != metadata.date.as_deref()
        {
            return Err(invalid("tracked move range metadata is inconsistent"));
        }
    } else {
        group.change = Some(Change {
            kind: ChangeKind::Move,
            id: metadata.id.clone(),
            author: metadata.author.clone(),
            date: metadata.date.clone(),
            name: Some(name.to_owned()),
        });
    }
    Ok(())
}

fn merge_move_wrapper_metadata(group: &mut Group, metadata: &Metadata) -> Result<()> {
    let change = group
        .change
        .as_ref()
        .ok_or_else(|| invalid("tracked move range has no owner metadata"))?;
    if change.author != metadata.author
        || metadata
            .date
            .as_deref()
            .is_some_and(|date| change.date.as_deref() != Some(date))
    {
        return Err(invalid("tracked move wrapper metadata is inconsistent"));
    }
    Ok(())
}

fn classify_wrapper(
    local: &[u8],
    namespace: NamespaceKind,
    parent: Option<&Frame>,
) -> Result<Option<(ChangeKind, Option<Side>)>> {
    if !namespace.is_word() {
        return Ok(None);
    }
    let kind = match local {
        b"ins" => Some((ChangeKind::Insert, None)),
        b"del" => Some((ChangeKind::Delete, None)),
        b"moveFrom" => Some((ChangeKind::Move, Some(Side::From))),
        b"moveTo" => Some((ChangeKind::Move, Some(Side::To))),
        _ => None,
    };
    let Some(kind) = kind else {
        return Ok(None);
    };
    if !legal_run_level_parent(parent) {
        return Err(invalid(
            "tracked revision wrapper has an unsupported paragraph or property owner",
        ));
    }
    Ok(Some(kind))
}

fn legal_run_level_parent(parent: Option<&Frame>) -> bool {
    parent.is_some_and(|parent| {
        parent.namespace.is_word()
            && matches!(
                parent.local.as_slice(),
                b"p" | b"hyperlink" | b"customXml" | b"smartTag" | b"dir" | b"bdo"
            )
    })
}

fn direct_run_track_change_frame(frames: &[Frame]) -> bool {
    frames.last().is_some_and(|frame| {
        frame.wrapper.is_some()
            && frame.namespace.is_word()
            && is_run_track_change_name(frame.local.as_slice())
    })
}

fn reject_run_parent_text(frames: &[Frame], range_active: bool, text: &[u8]) -> Result<()> {
    if text.iter().all(u8::is_ascii_whitespace) {
        return Ok(());
    }
    if (frames.iter().any(|frame| frame.wrapper.is_some()) || range_active)
        && frames.last().is_some_and(|frame| {
            frame.namespace.is_word() && matches!(frame.local.as_slice(), b"r" | b"rPr")
        })
    {
        return Err(invalid(
            "tracked revision run content has character data outside a text element",
        ));
    }
    Ok(())
}

fn is_run_track_change_name(local: &[u8]) -> bool {
    matches!(local, b"ins" | b"del" | b"moveFrom" | b"moveTo")
}

fn validate_run_track_change_child(
    parent: Option<&Frame>,
    local: &[u8],
    namespace: NamespaceKind,
) -> Result<()> {
    let Some(parent) = parent else {
        return Ok(());
    };
    if !(parent.wrapper.is_some()
        && parent.namespace.is_word()
        && is_run_track_change_name(parent.local.as_slice()))
    {
        return Ok(());
    }
    if namespace == NamespaceKind::Unbound {
        return Err(invalid(
            "tracked revision wrapper has an unqualified run-content child",
        ));
    }
    if namespace.is_word() && !is_run_track_change_child(local) {
        return Err(invalid(
            "tracked revision wrapper has a child outside CT_RunTrackChange run content",
        ));
    }
    Ok(())
}

fn validate_run_child(
    frames: &mut [Frame],
    range_active: bool,
    local: &[u8],
    namespace: NamespaceKind,
) -> Result<()> {
    let active_owner = frames.iter().any(|frame| frame.wrapper.is_some()) || range_active;
    let Some(parent) = frames.last_mut() else {
        return Ok(());
    };
    if !(active_owner && parent.namespace.is_word() && parent.local.as_slice() == b"r") {
        return Ok(());
    }
    if namespace == NamespaceKind::Unbound {
        return Err(invalid("tracked revision run has an unqualified child"));
    }
    if namespace.is_word() && !is_run_inner_content_child(local) {
        return Err(invalid(
            "tracked revision run has a child outside CT_R run content",
        ));
    }
    if namespace.is_word() && local == b"rPr" {
        if parent.run_rpr_seen {
            return Err(invalid("tracked revision run has duplicate rPr"));
        }
        if parent.run_content_started {
            return Err(invalid("tracked revision run has rPr after run content"));
        }
        parent.run_rpr_seen = true;
    } else {
        parent.run_content_started = true;
    }
    Ok(())
}

fn is_run_inner_content_child(local: &[u8]) -> bool {
    matches!(
        local,
        b"rPr"
            | b"br"
            | b"t"
            | b"contentPart"
            | b"delText"
            | b"instrText"
            | b"delInstrText"
            | b"noBreakHyphen"
            | b"softHyphen"
            | b"dayShort"
            | b"monthShort"
            | b"yearShort"
            | b"dayLong"
            | b"monthLong"
            | b"yearLong"
            | b"annotationRef"
            | b"footnoteRef"
            | b"endnoteRef"
            | b"separator"
            | b"continuationSeparator"
            | b"sym"
            | b"pgNum"
            | b"cr"
            | b"tab"
            | b"object"
            | b"fldChar"
            | b"ruby"
            | b"footnoteReference"
            | b"endnoteReference"
            | b"commentReference"
            | b"drawing"
            | b"ptab"
            | b"lastRenderedPageBreak"
    )
}

fn validate_run_properties_child(
    frames: &mut [Frame],
    range_active: bool,
    local: &[u8],
    namespace: NamespaceKind,
) -> Result<()> {
    let active_owner = frames.iter().any(|frame| frame.wrapper.is_some()) || range_active;
    let Some(parent) = frames.last_mut() else {
        return Ok(());
    };
    if !(active_owner && parent.namespace.is_word() && parent.local.as_slice() == b"rPr") {
        return Ok(());
    }
    if namespace == NamespaceKind::Unbound {
        return Err(invalid(
            "tracked revision run properties have an unqualified child",
        ));
    }
    if namespace.is_word() && !is_run_properties_child(local) {
        return Err(invalid(
            "tracked revision run properties have a child outside CT_RPr",
        ));
    }
    if namespace.is_word() {
        if local == b"rPrChange" {
            if parent.run_properties_change_seen {
                return Err(invalid(
                    "tracked revision run properties have duplicate rPrChange",
                ));
            }
            parent.run_properties_change_seen = true;
        } else if let Some(bit) = run_property_bit(local) {
            if parent.run_properties_change_seen {
                return Err(invalid(
                    "tracked revision run properties have a base child after rPrChange",
                ));
            }
            if parent.run_properties_seen & bit != 0 {
                return Err(invalid(
                    "tracked revision run properties have a duplicate base child",
                ));
            }
            parent.run_properties_seen |= bit;
        }
    }
    Ok(())
}

fn is_run_properties_child(local: &[u8]) -> bool {
    local == b"rPrChange" || run_property_bit(local).is_some()
}

fn run_property_bit(local: &[u8]) -> Option<u64> {
    let bit = match local {
        b"rStyle" => 0,
        b"rFonts" => 1,
        b"b" => 2,
        b"bCs" => 3,
        b"i" => 4,
        b"iCs" => 5,
        b"caps" => 6,
        b"smallCaps" => 7,
        b"strike" => 8,
        b"dstrike" => 9,
        b"outline" => 10,
        b"shadow" => 11,
        b"emboss" => 12,
        b"imprint" => 13,
        b"noProof" => 14,
        b"snapToGrid" => 15,
        b"vanish" => 16,
        b"webHidden" => 17,
        b"color" => 18,
        b"spacing" => 19,
        b"w" => 20,
        b"kern" => 21,
        b"position" => 22,
        b"sz" => 23,
        b"szCs" => 24,
        b"highlight" => 25,
        b"u" => 26,
        b"effect" => 27,
        b"bdr" => 28,
        b"shd" => 29,
        b"fitText" => 30,
        b"vertAlign" => 31,
        b"rtl" => 32,
        b"cs" => 33,
        b"em" => 34,
        b"lang" => 35,
        b"eastAsianLayout" => 36,
        b"specVanish" => 37,
        b"oMath" => 38,
        _ => return None,
    };
    Some(1u64 << bit)
}

fn reject_active_markup_compatibility(
    namespace: NamespaceKind,
    frames: &[Frame],
    open_ranges: &HashMap<(Side, String), OpenRange>,
) -> Result<()> {
    if namespace == NamespaceKind::MarkupCompatibility
        && (frames.iter().any(|frame| frame.wrapper.is_some()) || !open_ranges.is_empty())
    {
        return Err(invalid(
            "tracked revisions inside markup-compatibility alternatives are unsupported",
        ));
    }
    Ok(())
}

fn is_run_track_change_child(local: &[u8]) -> bool {
    matches!(
        local,
        b"customXml"
            | b"smartTag"
            | b"sdt"
            | b"dir"
            | b"bdo"
            | b"r"
            | b"proofErr"
            | b"permStart"
            | b"permEnd"
            | b"bookmarkStart"
            | b"bookmarkEnd"
            | b"moveFromRangeStart"
            | b"moveFromRangeEnd"
            | b"moveToRangeStart"
            | b"moveToRangeEnd"
            | b"commentRangeStart"
            | b"commentRangeEnd"
            | b"customXmlInsRangeStart"
            | b"customXmlInsRangeEnd"
            | b"customXmlDelRangeStart"
            | b"customXmlDelRangeEnd"
            | b"customXmlMoveFromRangeStart"
            | b"customXmlMoveFromRangeEnd"
            | b"customXmlMoveToRangeStart"
            | b"customXmlMoveToRangeEnd"
            | b"ins"
            | b"del"
            | b"moveFrom"
            | b"moveTo"
    )
}

fn classify_range(local: &[u8], namespace: NamespaceKind) -> Option<(Side, bool)> {
    if !namespace.is_word() {
        return None;
    }
    match local {
        b"moveFromRangeStart" => Some((Side::From, true)),
        b"moveFromRangeEnd" => Some((Side::From, false)),
        b"moveToRangeStart" => Some((Side::To, true)),
        b"moveToRangeEnd" => Some((Side::To, false)),
        _ => None,
    }
}

fn is_custom_xml_range_marker(local: &[u8], namespace: NamespaceKind) -> bool {
    namespace.is_word()
        && matches!(
            local,
            b"customXmlInsRangeStart"
                | b"customXmlInsRangeEnd"
                | b"customXmlDelRangeStart"
                | b"customXmlDelRangeEnd"
                | b"customXmlMoveFromRangeStart"
                | b"customXmlMoveFromRangeEnd"
                | b"customXmlMoveToRangeStart"
                | b"customXmlMoveToRangeEnd"
        )
}

fn unsupported_property_revision(local: &[u8]) -> bool {
    matches!(
        local,
        b"pPrChange"
            | b"rPrChange"
            | b"sectPrChange"
            | b"tblGridChange"
            | b"tblPrChange"
            | b"tblPrExChange"
            | b"tcPrChange"
            | b"trPrChange"
            | b"numberingChange"
            | b"cellIns"
            | b"cellDel"
            | b"cellMerge"
    )
}

fn unsupported_dependency(local: &[u8]) -> bool {
    matches!(
        local,
        b"p" | b"tbl"
            | b"tr"
            | b"tc"
            | b"tblPr"
            | b"trPr"
            | b"tcPr"
            | b"hyperlink"
            | b"customXml"
            | b"sdt"
            | b"sdtContent"
            | b"fldSimple"
            | b"fldChar"
            | b"instrText"
            | b"delInstrText"
            | b"bookmarkStart"
            | b"bookmarkEnd"
            | b"commentRangeStart"
            | b"commentRangeEnd"
            | b"permStart"
            | b"permEnd"
            | b"customXmlInsRangeStart"
            | b"customXmlInsRangeEnd"
            | b"customXmlDelRangeStart"
            | b"customXmlDelRangeEnd"
            | b"customXmlMoveFromRangeStart"
            | b"customXmlMoveFromRangeEnd"
            | b"customXmlMoveToRangeStart"
            | b"customXmlMoveToRangeEnd"
            | b"commentReference"
            | b"footnoteReference"
            | b"endnoteReference"
            | b"contentPart"
            | b"drawing"
            | b"pict"
            | b"object"
            | b"altChunk"
            | b"subDoc"
            | b"sectPr"
    )
}

fn has_field_or_control_ancestor(frames: &[Frame]) -> bool {
    frames.iter().any(|frame| {
        frame.namespace.is_word()
            && matches!(
                frame.local.as_slice(),
                b"fldSimple" | b"sdt" | b"sdtContent"
            )
    })
}

fn parse_metadata(
    source: &[u8],
    start: usize,
    end: usize,
    element: &quick_xml::events::BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
    total: &mut usize,
    require_move_name_and_date: bool,
) -> Result<Metadata> {
    let mut id = None;
    let mut author = None;
    let mut date = None;
    let mut name = None;
    let mut displaced_by_custom_xml = false;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_word_namespace(&namespace) {
            continue;
        }
        match attribute.key.local_name().as_ref() {
            b"id" => {
                if id.is_some() {
                    return Err(invalid("tracked revision has duplicate w:id"));
                }
                let value = decode_attribute(&attribute, decoder)?;
                validate_integer(&value)?;
                charge_metadata(total, value.len(), limits)?;
                id = Some(value);
            },
            b"author" => {
                if author.is_some() {
                    return Err(invalid("tracked revision has duplicate w:author"));
                }
                let value = decode_attribute(&attribute, decoder)?;
                charge_metadata(total, value.len(), limits)?;
                author = Some(value);
            },
            b"date" => {
                if date.is_some() {
                    return Err(invalid("tracked revision has duplicate w:date"));
                }
                let value = decode_attribute(&attribute, decoder)?;
                DateTime::new(value.clone())
                    .map_err(|_| invalid("tracked revision w:date is not a valid xsd:dateTime"))?;
                charge_metadata(total, value.len(), limits)?;
                date = Some(value);
            },
            b"name" if require_move_name_and_date => {
                if name.is_some() {
                    return Err(invalid("tracked move range has duplicate w:name"));
                }
                let value = decode_attribute(&attribute, decoder)?;
                charge_metadata(total, value.len(), limits)?;
                name = Some(value);
            },
            b"name" => {
                return Err(invalid(
                    "w:name is only valid on a tracked move range start",
                ));
            },
            b"displacedByCustomXml" if require_move_name_and_date => {
                if displaced_by_custom_xml {
                    return Err(invalid(
                        "tracked move range has duplicate w:displacedByCustomXml",
                    ));
                }
                let value = decode_attribute(&attribute, decoder)?;
                validate_displaced_by_custom_xml(&value)?;
                charge_metadata(total, value.len(), limits)?;
                displaced_by_custom_xml = true;
            },
            b"displacedByCustomXml" => {
                return Err(invalid(
                    "w:displacedByCustomXml is only valid on a tracked move range marker",
                ));
            },
            _ => {
                return Err(invalid(
                    "tracked revision uses an unsupported WordprocessingML attribute",
                ));
            },
        }
    }
    let id = id.ok_or_else(|| invalid("tracked revision requires w:id"))?;
    let author = author.ok_or_else(|| invalid("tracked revision requires w:author"))?;
    if require_move_name_and_date && date.is_none() {
        return Err(invalid("tracked move range start requires w:date"));
    }
    if require_move_name_and_date && name.is_none() {
        return Err(invalid("tracked move range start requires w:name"));
    }
    let key = normalize_integer(&id)?;
    let _ = (source, start, end);
    Ok(Metadata {
        key,
        id,
        author,
        date,
        name,
    })
}

fn parse_range_end_metadata(
    element: &quick_xml::events::BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
    total: &mut usize,
) -> Result<Metadata> {
    let mut id = None;
    let mut displaced_by_custom_xml = false;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_word_namespace(&namespace) {
            continue;
        }
        if attribute.key.local_name().as_ref() == b"id" {
            if id.is_some() {
                return Err(invalid("tracked move range end has duplicate w:id"));
            }
            let value = decode_attribute(&attribute, decoder)?;
            validate_integer(&value)?;
            charge_metadata(total, value.len(), limits)?;
            id = Some(value);
        } else if attribute.key.local_name().as_ref() == b"displacedByCustomXml" {
            if displaced_by_custom_xml {
                return Err(invalid(
                    "tracked move range end has duplicate w:displacedByCustomXml",
                ));
            }
            let value = decode_attribute(&attribute, decoder)?;
            validate_displaced_by_custom_xml(&value)?;
            charge_metadata(total, value.len(), limits)?;
            displaced_by_custom_xml = true;
        } else {
            return Err(invalid(
                "tracked move range end has an unsupported WordprocessingML attribute",
            ));
        }
    }
    let id = id.ok_or_else(|| invalid("tracked move range end requires w:id"))?;
    let key = normalize_integer(&id)?;
    Ok(Metadata {
        key,
        id,
        author: String::new(),
        date: None,
        name: None,
    })
}

fn validate_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
    _total: &mut usize,
    _metadata: bool,
) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("tracked revision attribute counter overflow"))?;
        if count > limits.max_attributes {
            return Err(limit_error("attributes", count, limits.max_attributes));
        }
        if attribute.value.len() > limits.max_attribute_bytes {
            return Err(limit_error(
                "attribute bytes",
                attribute.value.len(),
                limits.max_attribute_bytes,
            ));
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_word_namespace(&namespace) && local.as_ref() == b"displacedByCustomXml" {
            let value = decode_attribute(&attribute, decoder)?;
            validate_displaced_by_custom_xml(&value)?;
        }
        if !is_namespace_declaration_name(attribute.key.as_ref())
            && matches!(namespace, ResolveResult::Unknown(_))
        {
            return Err(invalid(
                "tracked revision XML uses an unbound attribute prefix",
            ));
        }
    }
    Ok(())
}

fn collect_hoisted_attributes(
    source: &[u8],
    start: usize,
    end: usize,
    element: &quick_xml::events::BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<(Vec<NamespaceDeclaration>, Vec<ScopeAttribute>)> {
    let tag = source
        .get(start..end)
        .ok_or_else(|| invalid("tracked revision opening tag is outside its source"))?;
    let raw_attributes = raw_tag_attributes(tag, limits)?;
    let mut attributes = element.attributes();
    let mut namespace_declarations = Vec::new();
    let mut scope_attributes = Vec::new();
    for raw in raw_attributes {
        let attribute = attributes
            .next()
            .ok_or_else(|| invalid("tracked revision XML attribute scan is inconsistent"))?
            .map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() != raw.name.as_slice() {
            return Err(invalid(
                "tracked revision XML attribute scan is inconsistent",
            ));
        }
        if is_namespace_declaration_name(&raw.name) {
            if namespace_declarations
                .iter()
                .any(|current: &NamespaceDeclaration| current.name == raw.name)
            {
                return Err(invalid(
                    "tracked revision has duplicate namespace declarations",
                ));
            }
            push_bounded(
                &mut namespace_declarations,
                NamespaceDeclaration {
                    name: raw.name,
                    value: raw.value,
                    raw: raw.raw,
                },
                "tracked revision namespace scope",
            )?;
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        let Some(namespace) = (match namespace {
            ResolveResult::Bound(Namespace(uri))
                if uri == MC
                    && matches!(
                        local.as_ref(),
                        b"Ignorable" | b"ProcessContent" | b"MustUnderstand"
                    ) =>
            {
                Some(MC)
            },
            ResolveResult::Bound(Namespace(uri))
                if uri == XML && matches!(local.as_ref(), b"space" | b"lang") =>
            {
                Some(XML)
            },
            _ => None,
        }) else {
            continue;
        };
        if scope_attributes.iter().any(|current: &ScopeAttribute| {
            current.namespace.as_slice() == namespace && current.local == local.as_ref()
        }) {
            return Err(invalid("tracked revision has duplicate scope attributes"));
        }
        let local = local.as_ref().to_vec();
        push_bounded(
            &mut scope_attributes,
            ScopeAttribute {
                local,
                namespace: namespace.to_vec(),
                value: raw.value,
                raw: raw.raw,
            },
            "tracked revision namespace scope",
        )?;
    }
    if attributes.next().is_some() {
        return Err(invalid(
            "tracked revision XML attribute scan is inconsistent",
        ));
    }
    let _ = decoder;
    Ok((namespace_declarations, scope_attributes))
}

fn raw_tag_name_end(tag: &[u8]) -> Result<usize> {
    if tag.first() != Some(&b'<') || tag.get(1) == Some(&b'/') {
        return Err(invalid("tracked revision opening tag is malformed"));
    }
    let mut cursor = 1usize;
    while tag
        .get(cursor)
        .is_some_and(|byte| !byte.is_ascii_whitespace() && !matches!(byte, b'>' | b'/'))
    {
        cursor += 1;
    }
    if cursor == 1 {
        return Err(invalid("tracked revision opening tag has no name"));
    }
    Ok(cursor)
}

fn is_namespace_declaration_name(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn raw_tag_attributes(tag: &[u8], limits: Limits) -> Result<Vec<RawAttribute>> {
    let mut cursor = raw_tag_name_end(tag)?;
    let mut output = Vec::new();
    while cursor < tag.len() {
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if cursor >= tag.len() || matches!(tag[cursor], b'>' | b'/') {
            break;
        }
        let name_start = cursor;
        while tag
            .get(cursor)
            .is_some_and(|byte| !byte.is_ascii_whitespace() && !matches!(byte, b'=' | b'>' | b'/'))
        {
            cursor += 1;
        }
        let name = tag
            .get(name_start..cursor)
            .filter(|name| !name.is_empty())
            .ok_or_else(|| invalid("tracked revision attribute has no name"))?;
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return Err(invalid("tracked revision attribute has no equals sign"));
        }
        cursor += 1;
        while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .ok_or_else(|| invalid("tracked revision attribute has no quoted value"))?;
        if !matches!(quote, b'"' | b'\'') {
            return Err(invalid("tracked revision attribute value is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while tag.get(cursor).is_some_and(|byte| *byte != quote) {
            cursor += 1;
        }
        let value = tag
            .get(value_start..cursor)
            .ok_or_else(|| invalid("tracked revision attribute value is unterminated"))?;
        if value.len() > limits.max_attribute_bytes {
            return Err(limit_error(
                "attribute bytes",
                value.len(),
                limits.max_attribute_bytes,
            ));
        }
        if cursor >= tag.len() {
            return Err(invalid("tracked revision attribute value is unterminated"));
        }
        cursor += 1;
        let raw = tag
            .get(name_start..cursor)
            .ok_or_else(|| invalid("tracked revision attribute span is invalid"))?;
        push_bounded(
            &mut output,
            RawAttribute {
                name: name.to_vec(),
                value: value.to_vec(),
                raw: raw.to_vec(),
            },
            "tracked revision raw attributes",
        )?;
        if output.len() > limits.max_attributes {
            return Err(limit_error(
                "attributes",
                output.len(),
                limits.max_attributes,
            ));
        }
    }
    Ok(output)
}

fn decode_attribute(
    attribute: &quick_xml::events::attributes::Attribute<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<String> {
    attribute
        .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
        .map(|value| value.into_owned())
        .map_err(|error| Error::Xml(error.to_string()))
}

fn validate_displaced_by_custom_xml(value: &str) -> Result<()> {
    if matches!(value, "next" | "prev") {
        Ok(())
    } else {
        Err(invalid(
            "w:displacedByCustomXml must be either next or prev",
        ))
    }
}

fn validate_integer(value: &str) -> Result<()> {
    let normalized = collapse_xml_whitespace(value);
    let digits = normalized
        .strip_prefix('+')
        .or_else(|| normalized.strip_prefix('-'))
        .unwrap_or(&normalized);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(invalid("tracked revision ID is not an xsd:integer"));
    }
    Ok(())
}

fn normalize_integer(value: &str) -> Result<String> {
    validate_integer(value)?;
    let normalized = collapse_xml_whitespace(value);
    let negative = normalized.starts_with('-');
    let digits = normalized
        .strip_prefix('+')
        .or_else(|| normalized.strip_prefix('-'))
        .unwrap_or(&normalized);
    let digits = digits.trim_start_matches('0');
    let digits = if digits.is_empty() { "0" } else { digits };
    if negative && digits != "0" {
        Ok(format!("-{digits}"))
    } else {
        Ok(digits.to_owned())
    }
}

fn collapse_xml_whitespace(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    let mut pending = false;
    for character in value.chars() {
        if matches!(character, '\u{9}' | '\u{A}' | '\u{D}' | ' ') {
            pending = true;
        } else {
            if pending && !output.is_empty() {
                output.push(' ');
            }
            output.push(character);
            pending = false;
        }
    }
    output
}

fn charge_metadata(total: &mut usize, added: usize, limits: Limits) -> Result<()> {
    *total = total
        .checked_add(added)
        .ok_or_else(|| limit_error("metadata bytes", usize::MAX, limits.max_metadata_bytes))?;
    if *total > limits.max_metadata_bytes {
        return Err(limit_error(
            "metadata bytes",
            *total,
            limits.max_metadata_bytes,
        ));
    }
    Ok(())
}

fn name_span(source: &[u8], start: usize, end: usize, closing: bool) -> Result<NameSpan> {
    let tag = source
        .get(start..end)
        .ok_or_else(|| invalid("tracked revision element span is outside its source"))?;
    let mut cursor = if closing { 2 } else { 1 };
    while tag.get(cursor).is_some_and(u8::is_ascii_whitespace) {
        cursor += 1;
    }
    let name_start = cursor;
    while tag
        .get(cursor)
        .is_some_and(|byte| !byte.is_ascii_whitespace() && *byte != b'>' && *byte != b'/')
    {
        cursor += 1;
    }
    let raw = tag
        .get(name_start..cursor)
        .ok_or_else(|| invalid("tracked revision element name is missing"))?;
    let local_start = raw
        .iter()
        .rposition(|byte| *byte == b':')
        .map_or(name_start, |offset| name_start + offset + 1);
    let local = &raw[local_start - name_start..];
    let replacement = if local == b"delText" {
        let mut replacement = Vec::new();
        replacement
            .try_reserve(1)
            .map_err(|_| Error::RevisionAllocation {
                resource: "tracked revision element name",
            })?;
        replacement.push(b't');
        replacement
    } else {
        return Err(invalid("tracked revision text element is not w:delText"));
    };
    Ok(NameSpan {
        span: Span::new(start + local_start, start + cursor)?,
        replacement,
    })
}

fn plan_group(
    source: &[u8],
    group: &Group,
    disposition: Disposition,
    splices: &mut Vec<Splice>,
) -> Result<()> {
    let kind = group
        .kind
        .ok_or_else(|| invalid("tracked revision group has no operation kind"))?;
    match kind {
        ChangeKind::Insert | ChangeKind::Delete => {
            let remove = matches!(
                (kind, disposition),
                (ChangeKind::Insert, Disposition::Reject)
                    | (ChangeKind::Delete, Disposition::Accept)
            );
            for wrapper in &group.wrappers {
                if remove {
                    push_bounded(
                        splices,
                        Splice {
                            span: wrapper.whole,
                            replacement: Vec::new(),
                        },
                        "tracked revision source splices",
                    )?;
                } else {
                    strip_wrapper(source, wrapper, splices)?;
                    if kind == ChangeKind::Delete {
                        for name in &wrapper.deleted_names {
                            push_bounded(
                                splices,
                                Splice {
                                    span: name.span,
                                    replacement: name.replacement.clone(),
                                },
                                "tracked revision source splices",
                            )?;
                        }
                    }
                }
            }
        },
        ChangeKind::Move => {
            let keep_from = disposition == Disposition::Reject;
            let keep_to = disposition == Disposition::Accept;
            plan_move_side(source, group, Side::From, keep_from, splices)?;
            plan_move_side(source, group, Side::To, keep_to, splices)?;
        },
    }
    Ok(())
}

fn plan_move_side(
    source: &[u8],
    group: &Group,
    side: Side,
    keep: bool,
    splices: &mut Vec<Splice>,
) -> Result<()> {
    if let Some(range) = match side {
        Side::From => group.from_range.as_ref(),
        Side::To => group.to_range.as_ref(),
    } {
        if keep {
            push_bounded(
                splices,
                Splice {
                    span: range.start_marker,
                    replacement: Vec::new(),
                },
                "tracked revision source splices",
            )?;
            push_bounded(
                splices,
                Splice {
                    span: range.end_marker,
                    replacement: Vec::new(),
                },
                "tracked revision source splices",
            )?;
        } else {
            // The range markers establish the move container; they do not
            // own every byte between them. Only the move wrapper below is
            // removed, so ordinary runs and opaque siblings survive.
            push_bounded(
                splices,
                Splice {
                    span: range.start_marker,
                    replacement: Vec::new(),
                },
                "tracked revision source splices",
            )?;
            push_bounded(
                splices,
                Splice {
                    span: range.end_marker,
                    replacement: Vec::new(),
                },
                "tracked revision source splices",
            )?;
        }
    }
    for wrapper in group
        .wrappers
        .iter()
        .filter(|wrapper| wrapper.side == Some(side))
    {
        if keep {
            strip_wrapper(source, wrapper, splices)?;
            if wrapper.kind == ChangeKind::Move && side == Side::From {
                for name in &wrapper.deleted_names {
                    push_bounded(
                        splices,
                        Splice {
                            span: name.span,
                            replacement: name.replacement.clone(),
                        },
                        "tracked revision source splices",
                    )?;
                }
            }
        } else {
            push_bounded(
                splices,
                Splice {
                    span: wrapper.whole,
                    replacement: Vec::new(),
                },
                "tracked revision source splices",
            )?;
        }
    }
    Ok(())
}

fn strip_wrapper(source: &[u8], wrapper: &Wrapper, splices: &mut Vec<Splice>) -> Result<()> {
    for child in &wrapper.children {
        let opening = source
            .get(child.span.start..child.span.end)
            .ok_or_else(|| invalid("tracked revision child opening is outside its source"))?;
        let name_end = raw_tag_name_end(opening)?;
        let mut insertion = Vec::new();
        for declaration in &wrapper.namespace_declarations {
            if let Some(existing) = child
                .namespace_declarations
                .iter()
                .find(|current| current.name == declaration.name)
            {
                if existing.value != declaration.value {
                    return Err(invalid(
                        "tracked revision child redeclares a wrapper namespace differently",
                    ));
                }
                continue;
            }
            insertion
                .try_reserve(1usize.saturating_add(declaration.raw.len()))
                .map_err(|_| Error::RevisionAllocation {
                    resource: "tracked revision namespace scope",
                })?;
            insertion.push(b' ');
            insertion.extend_from_slice(&declaration.raw);
        }
        for attribute in &wrapper.scope_attributes {
            if let Some(existing) = child.scope_attributes.iter().find(|current| {
                current.namespace == attribute.namespace && current.local == attribute.local
            }) {
                if existing.value != attribute.value {
                    return Err(invalid(
                        "tracked revision child redeclares a wrapper scope attribute differently",
                    ));
                }
                continue;
            }
            insertion
                .try_reserve(1usize.saturating_add(attribute.raw.len()))
                .map_err(|_| Error::RevisionAllocation {
                    resource: "tracked revision namespace scope",
                })?;
            insertion.push(b' ');
            insertion.extend_from_slice(&attribute.raw);
        }
        if !insertion.is_empty() {
            let insertion_start = child
                .span
                .start
                .checked_add(name_end)
                .ok_or_else(|| invalid("tracked revision namespace scope offset overflow"))?;
            push_bounded(
                splices,
                Splice {
                    span: Span::new(insertion_start, insertion_start)?,
                    replacement: insertion,
                },
                "tracked revision source splices",
            )?;
        }
    }
    push_bounded(
        splices,
        Splice {
            span: wrapper.start,
            replacement: Vec::new(),
        },
        "tracked revision source splices",
    )?;
    if let Some(end) = wrapper.end {
        push_bounded(
            splices,
            Splice {
                span: end,
                replacement: Vec::new(),
            },
            "tracked revision source splices",
        )?;
    }
    Ok(())
}

#[derive(Debug, Clone)]
struct Splice {
    span: Span,
    replacement: Vec<u8>,
}

fn apply_splices(source: &[u8], splices: &mut [Splice], limit: usize) -> Result<Vec<u8>> {
    splices.sort_unstable_by_key(|splice| (splice.span.start, splice.span.end));
    let mut cursor = 0usize;
    let mut removed = 0usize;
    let mut added = 0usize;
    for splice in splices.iter() {
        if splice.span.end > source.len() || splice.span.start < cursor {
            return Err(invalid(
                "tracked revision source edits overlap or are out of bounds",
            ));
        }
        cursor = splice.span.end;
        removed = removed
            .checked_add(splice.span.len())
            .ok_or_else(|| invalid("tracked revision output size overflow"))?;
        added = added
            .checked_add(splice.replacement.len())
            .ok_or_else(|| invalid("tracked revision output size overflow"))?;
    }
    let output_len = source
        .len()
        .checked_sub(removed)
        .and_then(|value| value.checked_add(added))
        .ok_or_else(|| invalid("tracked revision output size overflow"))?;
    if output_len > limit {
        return Err(limit_error("output bytes", output_len, limit));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| Error::RevisionAllocation {
            resource: "tracked revision XML output",
        })?;
    cursor = 0;
    for splice in splices {
        output.extend_from_slice(&source[cursor..splice.span.start]);
        output.extend_from_slice(&splice.replacement);
        cursor = splice.span.end;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == W_TRANSITIONAL || *uri == W_STRICT => {
            NamespaceKind::Word
        },
        ResolveResult::Bound(Namespace(uri)) if *uri == MC => NamespaceKind::MarkupCompatibility,
        ResolveResult::Bound(_) | ResolveResult::Unknown(_) => NamespaceKind::Other,
        ResolveResult::Unbound => NamespaceKind::Unbound,
    }
}

fn reject_unknown_namespace(namespace: &ResolveResult<'_>) -> Result<()> {
    if matches!(namespace, ResolveResult::Unknown(_)) {
        return Err(invalid(
            "tracked revision XML uses an unbound element prefix",
        ));
    }
    Ok(())
}

fn is_word_namespace(namespace: &ResolveResult<'_>) -> bool {
    namespace_kind(namespace) == NamespaceKind::Word
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("tracked revision XML offset does not fit usize"))
}

fn limit_error(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::RevisionLimit {
        resource,
        actual,
        maximum,
    }
}

fn package_story_limits(limits: Limits) -> StoryLimits {
    let defaults = StoryLimits::default();
    StoryLimits {
        max_story_bytes: limits.max_source_bytes,
        max_total_story_bytes: defaults.max_total_story_bytes.max(limits.max_source_bytes),
        ..defaults
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn push_bounded<T>(values: &mut Vec<T>, value: T, resource: &'static str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|_| Error::RevisionAllocation { resource })?;
    values.push(value);
    Ok(())
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use super::{ChangeKind, Error, Limits, Snapshot};
    use crate::Package;
    use crate::writer::{RevisionKind, RevisionMetadata};

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn document(body: &str) -> Vec<u8> {
        format!("<w:document xmlns:w=\"{W}\"><w:body>{body}</w:body></w:document>").into_bytes()
    }

    #[test]
    fn inline_dispositions_preserve_source_and_inverse_exactly() {
        let source = document(
            r#"<w:p><w:r><w:t>base</w:t></w:r><w:ins w:id="1" w:author="Alice"><w:r><w:t xml:space="preserve"> add </w:t></w:r></w:ins><w:del w:id="2" w:author="Bob"><w:r><w:delText xml:space="preserve"> old </w:delText></w:r></w:del></w:p>"#,
        );
        let snapshot = Snapshot::from_xml(source.clone()).expect("parse inline revisions");
        assert_eq!(snapshot.changes().len(), 2);
        assert_eq!(snapshot.changes()[0].kind(), ChangeKind::Insert);
        assert_eq!(snapshot.changes()[1].kind(), ChangeKind::Delete);

        let mut transaction = snapshot.edit();
        transaction
            .accept(0)
            .expect("stage insertion acceptance")
            .reject(1)
            .expect("stage deletion rejection");
        let commit = transaction.commit().expect("commit dispositions");
        assert!(commit.changed());
        assert_eq!(snapshot.source(), source.as_slice());
        assert!(
            !commit
                .snapshot()
                .source()
                .windows(b"w:ins".len())
                .any(|window| window == b"w:ins")
        );
        assert!(
            !commit
                .snapshot()
                .source()
                .windows(b"w:del".len())
                .any(|window| window == b"w:del")
        );
        assert!(
            commit
                .snapshot()
                .source()
                .windows(b" add ".len())
                .any(|window| window == b" add ")
        );
        assert!(
            commit
                .snapshot()
                .source()
                .windows(b" old ".len())
                .any(|window| window == b" old ")
        );
        assert!(
            commit
                .snapshot()
                .source()
                .windows(b"w:t".len())
                .any(|window| window == b"w:t")
        );

        let inverse = commit.patch().inverse();
        let restored = inverse
            .apply(commit.snapshot())
            .expect("apply exact inverse");
        assert_eq!(restored.source(), source.as_slice());
    }

    #[test]
    fn retained_payload_hoists_wrapper_namespace_and_scope_attributes() {
        let source = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice" xmlns:x="urn:opaque" xmlns:c="http://schemas.openxmlformats.org/markup-compatibility/2006" c:Ignorable="x" xml:space="preserve" xml:lang="en"><w:r x:vendor="a>b"><x:opaque>  <x:child/>  payload<![CDATA[ & ]]></x:opaque><w:t> x </w:t></w:r></w:ins></w:p>"#,
        );
        let snapshot = Snapshot::from_xml(source.clone()).expect("parse scoped payload");
        let mut transaction = snapshot.edit();
        transaction.accept(0).expect("accept insertion");
        let commit = transaction.commit().expect("commit scoped payload");
        let output = std::str::from_utf8(commit.snapshot().source()).unwrap();
        assert!(!output.contains("w:ins"));
        for expected in [
            r#"xmlns:x="urn:opaque""#,
            r#"xmlns:c="http://schemas.openxmlformats.org/markup-compatibility/2006""#,
            r#"c:Ignorable="x""#,
            r#"xml:space="preserve""#,
            r#"xml:lang="en""#,
            r#"x:vendor="a>b""#,
            "<x:opaque>  <x:child/>  payload<![CDATA[ & ]]></x:opaque>",
        ] {
            assert!(output.contains(expected), "missing {expected} in {output}");
        }
        let restored = commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .expect("inverse scoped payload");
        assert_eq!(restored.source(), source.as_slice());
    }

    #[test]
    fn deleted_text_rewrites_only_the_local_name_and_reopens_exact_xml() {
        let source = document(
            r#"<w:p><w:del w:id="1" w:author="Alice"><w:r><w:delText>old</w:delText></w:r></w:del></w:p>"#,
        );
        let snapshot = Snapshot::from_xml(source).expect("parse deleted text");
        let mut transaction = snapshot.edit();
        transaction.reject(0).expect("reject deletion");
        let commit = transaction.commit().expect("commit deletion rejection");
        let expected = document(r#"<w:p><w:r><w:t>old</w:t></w:r></w:p>"#);
        assert_eq!(commit.snapshot().source(), expected.as_slice());
        assert!(
            !commit
                .snapshot()
                .source()
                .windows(b"w:w:t".len())
                .any(|window| { window == b"w:w:t" })
        );
        Snapshot::from_xml(expected).expect("reopen exact accepted XML");
    }

    #[test]
    fn empty_scoped_wrapper_can_be_removed_without_a_child() {
        let source = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice" xmlns:x="urn:opaque" xml:space="preserve"/></w:p>"#,
        );
        let snapshot = Snapshot::from_xml(source).expect("parse empty scoped insertion");
        let mut transaction = snapshot.edit();
        transaction.accept(0).expect("accept empty insertion");
        let commit = transaction.commit().expect("remove empty insertion");
        let expected = document(r#"<w:p></w:p>"#);
        assert_eq!(commit.snapshot().source(), expected.as_slice());
        Snapshot::from_xml(expected).expect("reopen empty-wrapper result");
    }

    #[test]
    fn run_track_change_child_grammar_and_parent_ownership_are_checked() {
        for child in [
            r#"<w:t>invalid direct text</w:t>"#,
            r#"<w:rPr/>"#,
            r#"<w:unknownWord/>"#,
        ] {
            let source = document(&format!(
                r#"<w:p><w:ins w:id="1" w:author="Alice">{child}</w:ins></w:p>"#
            ));
            assert!(
                Snapshot::from_xml(source).is_err(),
                "direct wrapper child unexpectedly accepted: {child}"
            );
        }

        let nested_wrapper = document(
            r#"<w:p><w:r><w:ins w:id="1" w:author="Alice"><w:r><w:t>x</w:t></w:r></w:ins></w:r></w:p>"#,
        );
        assert!(Snapshot::from_xml(nested_wrapper).is_err());

        let range_in_run = document(
            r#"<w:p><w:r><w:moveFromRangeStart w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z" w:name="m"/></w:r></w:p>"#,
        );
        assert!(Snapshot::from_xml(range_in_run).is_err());
    }

    #[test]
    fn unqualified_and_nested_word_run_content_are_refused() {
        let unqualified_direct =
            document(r#"<w:p><w:ins w:id="1" w:author="Alice"><t>unqualified</t></w:ins></w:p>"#);
        assert!(Snapshot::from_xml(unqualified_direct).is_err());

        let unknown_nested_word = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:unknownWord/></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(unknown_nested_word).is_err());

        let unknown_run_properties = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:rPr><w:unknownWord/></w:rPr><w:t>x</w:t></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(unknown_run_properties).is_err());

        let valid_run_properties = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:rPr><w:b w:val="1"/></w:rPr><w:t>x</w:t></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(valid_run_properties).is_ok());

        let unqualified_nested = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><t>unqualified</t></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(unqualified_nested).is_err());

        let unknown_in_move_range = document(
            r#"<w:p><w:moveFromRangeStart w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z" w:name="m"/><w:r><w:unknownWord/></w:r><w:moveFromRangeEnd w:id="1"/><w:moveToRangeStart w:id="2" w:author="Alice" w:date="2026-01-01T00:00:00Z" w:name="m"/><w:r><w:t>x</w:t></w:r><w:moveToRangeEnd w:id="2"/></w:p>"#,
        );
        assert!(Snapshot::from_xml(unknown_in_move_range).is_err());

        let foreign_payload = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><x:opaque xmlns:x="urn:opaque"/></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(foreign_payload).is_ok());
    }

    #[test]
    fn markup_compatibility_payloads_are_refused_with_tracked_revisions() {
        let alternate_inside_wrapper = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback><w:r><w:t>x</w:t></w:r></mc:Fallback></mc:AlternateContent></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(alternate_inside_wrapper).is_err());

        let alternate_inside_run = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback><w:t>x</w:t></mc:Fallback></mc:AlternateContent></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(alternate_inside_run).is_err());

        let revision_inside_alternate = document(
            r#"<w:p><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="w"><w:ins w:id="1" w:author="Alice"><w:r><w:t>x</w:t></w:r></w:ins></mc:Choice><mc:Fallback><w:r><w:t>fallback</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>"#,
        );
        assert!(Snapshot::from_xml(revision_inside_alternate).is_err());
    }

    #[test]
    fn run_and_run_properties_character_data_is_rejected_but_whitespace_survives() {
        for content in [
            r#"<w:r>text</w:r>"#,
            r#"<w:r><![CDATA[text]]></w:r>"#,
            r#"<w:rPr>text</w:rPr>"#,
            r#"<w:rPr><![CDATA[text]]></w:rPr>"#,
        ] {
            let source = document(&format!(
                r#"<w:p><w:ins w:id="1" w:author="Alice">{content}</w:ins></w:p>"#
            ));
            assert!(
                Snapshot::from_xml(source).is_err(),
                "invalid character data unexpectedly accepted: {content}"
            );
        }

        let source = document(
            "<w:p><w:ins w:id=\"1\" w:author=\"Alice\"><w:r> \n <w:rPr>\t </w:rPr> \n <w:t>x</w:t>\n</w:r></w:ins></w:p>",
        );
        let snapshot = Snapshot::from_xml(source.clone()).expect("parse whitespace-only run");
        let mut transaction = snapshot.edit();
        transaction.accept(0).expect("accept insertion");
        let commit = transaction.commit().expect("remove insertion");
        let expected = document("<w:p><w:r> \n <w:rPr>\t </w:rPr> \n <w:t>x</w:t>\n</w:r></w:p>");
        assert_eq!(commit.snapshot().source(), expected.as_slice());
        Snapshot::from_xml(expected).expect("reopen whitespace-preserving output");
    }

    #[test]
    fn run_and_run_properties_cardinality_and_order_follow_schema() {
        let duplicate_rpr = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:rPr/><w:rPr/><w:t>x</w:t></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(duplicate_rpr).is_err());

        let late_rpr = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:t>x</w:t><w:rPr/></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(late_rpr).is_err());

        let duplicate_base_property = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:rPr><w:b/><w:b/></w:rPr><w:t>x</w:t></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(duplicate_base_property).is_err());

        let unordered_base_properties = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:rPr><w:i/><w:b/></w:rPr><w:t>x</w:t></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(unordered_base_properties).is_ok());
    }

    #[test]
    fn dependency_bearing_content_and_custom_xml_closures_are_refused() {
        let content_part = document(
            r#"<w:p><w:ins w:id="1" w:author="Alice"><w:r><w:contentPart xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(content_part).is_err());

        let custom_xml_closure = document(
            r#"<w:p><w:customXmlInsRangeStart w:id="9" w:author="Alice" w:date="2026-01-01T00:00:00Z"/><w:ins w:id="1" w:author="Alice"><w:r><w:t>x</w:t></w:r></w:ins><w:customXmlInsRangeEnd w:id="9"/></w:p>"#,
        );
        assert!(Snapshot::from_xml(custom_xml_closure).is_err());
    }

    #[test]
    fn move_dispositions_require_complete_ranges_and_are_reversible() {
        let source = document(
            r#"<w:p><w:moveFromRangeStart w:id="2" w:author="Alice" w:date="2026-01-01T00:00:00Z" w:name="move1"/><w:moveFrom w:id="3" w:author="Alice"><w:r><w:delText>old</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="2"/><w:moveToRangeStart w:id="0" w:author="Alice" w:date="2026-01-01T00:00:00Z" w:name="move1"/><w:moveTo w:id="1" w:author="Alice"><w:r><w:t>new</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="0"/></w:p>"#,
        );
        let snapshot = Snapshot::from_xml(source.clone()).expect("parse tracked move");
        assert_eq!(snapshot.changes().len(), 1);
        assert_eq!(snapshot.changes()[0].kind(), ChangeKind::Move);

        let mut accept = snapshot.edit();
        accept.accept(0).expect("stage move acceptance");
        let accepted = accept.commit().expect("accept move");
        let accepted_source = std::str::from_utf8(accepted.snapshot().source()).unwrap();
        assert!(!accepted_source.contains("moveFrom"));
        assert!(!accepted_source.contains("moveTo"));
        assert!(!accepted_source.contains("moveFromRange"));
        assert!(!accepted_source.contains("moveToRange"));
        assert!(!accepted_source.contains("old"));
        assert!(accepted_source.contains(">new<"));

        let restored = accepted
            .patch()
            .inverse()
            .apply(accepted.snapshot())
            .expect("restore accepted move");
        assert_eq!(restored.source(), source.as_slice());

        let mut reject = snapshot.edit();
        reject.reject(0).expect("stage move rejection");
        let rejected = reject.commit().expect("reject move");
        let rejected_source = std::str::from_utf8(rejected.snapshot().source()).unwrap();
        assert!(rejected_source.contains(">old<"));
        assert!(!rejected_source.contains("new"));
        assert!(!rejected_source.contains("moveFrom"));
        assert!(!rejected_source.contains("moveTo"));
    }

    #[test]
    fn move_ranges_pair_by_name_and_preserve_unmarked_content() {
        let source = document(
            r#"<w:p><w:moveToRangeStart w:id="0" w:author="Alice" w:date="2026-01-01T00:00:00Z" w:name="move1"/><w:r><w:t>to-ordinary</w:t></w:r><w:moveTo w:id="1" w:author="Alice"><w:r><w:t>new</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="0"/><w:moveFromRangeStart w:id="2" w:author="Alice" w:date="2026-01-01T00:00:00Z" w:name="move1"/><w:r><w:t>from-ordinary</w:t></w:r><w:moveFrom w:id="3" w:author="Alice"><w:r><w:delText>old</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="2"/></w:p>"#,
        );
        let snapshot = Snapshot::from_xml(source.clone()).expect("parse named move");
        assert_eq!(snapshot.changes().len(), 1);
        assert_eq!(snapshot.changes()[0].id(), "0");
        assert_eq!(snapshot.changes()[0].name(), Some("move1"));

        let mut accept = snapshot.edit();
        accept.accept(0).expect("stage named move acceptance");
        let accepted = accept.commit().expect("accept named move");
        let accepted_source = std::str::from_utf8(accepted.snapshot().source()).unwrap();
        for expected in ["to-ordinary", "from-ordinary", ">new<"] {
            assert!(accepted_source.contains(expected), "missing {expected}");
        }
        assert!(!accepted_source.contains(">old<"));
        let accepted_expected = document(
            r#"<w:p><w:r><w:t>to-ordinary</w:t></w:r><w:r><w:t>new</w:t></w:r><w:r><w:t>from-ordinary</w:t></w:r></w:p>"#,
        );
        assert_eq!(accepted.snapshot().source(), accepted_expected.as_slice());
        Snapshot::from_xml(accepted_expected).expect("reopen exact accepted move XML");
        assert_eq!(
            accepted
                .patch()
                .inverse()
                .apply(accepted.snapshot())
                .unwrap()
                .source(),
            source.as_slice()
        );

        let mut reject = snapshot.edit();
        reject.reject(0).expect("stage named move rejection");
        let rejected = reject.commit().expect("reject named move");
        let rejected_source = std::str::from_utf8(rejected.snapshot().source()).unwrap();
        for expected in ["to-ordinary", "from-ordinary", ">old<"] {
            assert!(rejected_source.contains(expected), "missing {expected}");
        }
        assert!(!rejected_source.contains(">new<"));
        let rejected_expected = document(
            r#"<w:p><w:r><w:t>to-ordinary</w:t></w:r><w:r><w:t>from-ordinary</w:t></w:r><w:r><w:t>old</w:t></w:r></w:p>"#,
        );
        assert_eq!(rejected.snapshot().source(), rejected_expected.as_slice());
        Snapshot::from_xml(rejected_expected).expect("reopen exact rejected move XML");
        assert_eq!(
            rejected
                .patch()
                .inverse()
                .apply(rejected.snapshot())
                .unwrap()
                .source(),
            source.as_slice()
        );
    }

    #[test]
    fn wrapper_only_moves_are_refused_without_complete_ranges() {
        let source = document(
            r#"<w:p><w:moveFrom w:id="4" w:author="Alice"><w:r><w:delText>old</w:delText></w:r></w:moveFrom><w:moveTo w:id="4" w:author="Alice"><w:r><w:t>new</w:t></w:r></w:moveTo></w:p>"#,
        );
        assert!(Snapshot::from_xml(source).is_err());
    }

    #[test]
    fn move_range_start_ids_are_unique_after_closed_ranges() {
        let source = document(
            r#"<w:p><w:moveFromRangeStart w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z" w:name="m"/><w:moveFrom w:id="2" w:author="A"><w:r><w:delText>x</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="1"/><w:moveToRangeStart w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z" w:name="m"/><w:moveTo w:id="3" w:author="A"><w:r><w:t>x</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="1"/></w:p>"#,
        );
        assert!(Snapshot::from_xml(source).is_err());
    }

    #[test]
    fn move_marker_attributes_are_schema_checked_and_invalid_names_are_refused() {
        let invalid_name = document(
            r#"<w:p><w:ins w:id="1" w:author="A" w:name="not-a-move"><w:r><w:t>x</w:t></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(invalid_name).is_err());

        let unknown_word_attribute = document(
            r#"<w:p><w:ins w:id="1" w:author="A" w:unsupported="x"><w:r><w:t>x</w:t></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(unknown_word_attribute).is_err());

        let invalid_displacement = document(
            r#"<w:p><w:moveFromRangeStart w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z" w:name="m" w:displacedByCustomXml="middle"/><w:moveFrom w:id="2" w:author="A"><w:r><w:delText>x</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="1"/><w:moveToRangeStart w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z" w:name="m"/><w:moveTo w:id="3" w:author="A"><w:r><w:t>x</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="2"/></w:p>"#,
        );
        assert!(Snapshot::from_xml(invalid_displacement).is_err());

        let valid_displacement = document(
            r#"<w:p><w:moveFromRangeStart w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z" w:name="m" w:displacedByCustomXml="next"/><w:moveFrom w:id="2" w:author="A"><w:r><w:delText>x</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="1" w:displacedByCustomXml="prev"/><w:moveToRangeStart w:id="3" w:author="A" w:date="2026-01-01T00:00:00Z" w:name="m" w:displacedByCustomXml="next"/><w:moveTo w:id="4" w:author="A"><w:r><w:t>x</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="3" w:displacedByCustomXml="prev"/></w:p>"#,
        );
        assert!(Snapshot::from_xml(valid_displacement).is_ok());
    }

    #[test]
    fn field_and_control_contained_revisions_are_refused() {
        let simple_field = document(
            r#"<w:p><w:fldSimple w:instr="PAGE"><w:ins w:id="1" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins></w:fldSimple></w:p>"#,
        );
        assert!(Snapshot::from_xml(simple_field).is_err());

        let control = document(
            r#"<w:p><w:sdt><w:sdtContent><w:ins w:id="1" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins></w:sdtContent></w:sdt></w:p>"#,
        );
        assert!(Snapshot::from_xml(control).is_err());

        let range_in_control = document(
            r#"<w:p><w:sdt><w:sdtContent><w:moveFromRangeStart w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z" w:name="m"/></w:sdtContent></w:sdt></w:p>"#,
        );
        assert!(Snapshot::from_xml(range_in_control).is_err());
    }

    #[test]
    fn unsupported_or_incomplete_owners_fail_before_transaction_state_changes() {
        let nested = document(
            r#"<w:p><w:ins w:id="1" w:author="A"><w:ins w:id="2" w:author="B"><w:r><w:t>x</w:t></w:r></w:ins></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(nested).is_err());

        let foreign_parent = document(
            r#"<x:p xmlns:x="urn:foreign"><w:ins w:id="1" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins></x:p>"#,
        );
        assert!(Snapshot::from_xml(foreign_parent).is_err());

        let incomplete_move = document(
            r#"<w:p><w:moveFrom w:id="1" w:author="A"><w:r><w:delText>x</w:delText></w:r></w:moveFrom></w:p>"#,
        );
        assert!(Snapshot::from_xml(incomplete_move).is_err());

        let missing_move_name = document(
            r#"<w:p><w:moveFromRangeStart w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"/><w:moveFromRangeEnd w:id="1"/><w:moveToRangeStart w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z" w:name="m"/><w:moveToRangeEnd w:id="2"/></w:p>"#,
        );
        assert!(Snapshot::from_xml(missing_move_name).is_err());

        let property =
            document(r#"<w:p><w:pPrChange w:id="1" w:author="A"><w:pPr/></w:pPrChange></w:p>"#);
        assert!(Snapshot::from_xml(property).is_err());

        let field = document(
            r#"<w:p><w:ins w:id="1" w:author="A"><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:ins></w:p>"#,
        );
        assert!(Snapshot::from_xml(field).is_err());

        let source = Snapshot::from_xml(document(
            r#"<w:p><w:ins w:id="1" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins></w:p>"#,
        ))
        .unwrap();
        let mut transaction = source.edit();
        assert!(transaction.accept(1).is_err());
        assert!(!transaction.is_changed());
    }

    #[test]
    fn no_op_and_limits_are_bounded_and_retryable() {
        let source_bytes =
            document(r#"<w:p><w:ins w:id="1" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins></w:p>"#);
        let source = Snapshot::from_xml(source_bytes.clone()).unwrap();
        let commit = source.edit().commit().expect("empty commit");
        assert!(!commit.changed());
        assert!(commit.patch().is_noop());
        let noop_target = commit.patch().apply(&source).expect("apply exact no-op");
        assert_eq!(noop_target.source(), source_bytes.as_slice());
        assert!(commit.patch().is_applied());
        assert!(commit.patch().apply(&source).is_err());

        let limits = Limits {
            max_events: 1,
            ..Limits::default()
        };
        let error = Snapshot::from_xml_with_limits(source_bytes, limits).unwrap_err();
        assert!(matches!(
            error,
            Error::RevisionLimit {
                resource: "events",
                ..
            }
        ));

        let change_limits = Limits {
            max_changes: 0,
            ..Limits::default()
        };
        let error = Snapshot::from_xml_with_limits(
            document(r#"<w:p><w:ins w:id="1" w:author="A"><w:r/></w:ins></w:p>"#),
            change_limits,
        )
        .unwrap_err();
        assert!(matches!(
            error,
            Error::RevisionLimit {
                resource: "changes",
                ..
            }
        ));

        let metadata_limits = Limits {
            max_metadata_bytes: 2,
            ..Limits::default()
        };
        let error = Snapshot::from_xml_with_limits(
            document(r#"<w:p><w:ins w:id="1" w:author="Alice"><w:r/></w:ins></w:p>"#),
            metadata_limits,
        )
        .unwrap_err();
        assert!(matches!(
            error,
            Error::RevisionLimit {
                resource: "metadata bytes",
                ..
            }
        ));
    }

    #[test]
    fn package_owner_binds_story_topology_and_publishes_inverse() {
        let mut package = Package::new().expect("new package");
        let metadata = RevisionMetadata::new("1", "Alice").expect("revision metadata");
        package
            .document_mut()
            .expect("document writer")
            .add_paragraph()
            .add_revision(RevisionKind::Insert, metadata)
            .add_run_with_text("new");
        let mut bytes = Cursor::new(Vec::new());
        package
            .to_plain_stream(&mut bytes)
            .expect("materialize package");
        let mut package =
            Package::from_reader(Cursor::new(bytes.into_inner())).expect("open package");
        let snapshot = package.tracked_revisions().expect("read package revisions");
        let mut transaction = snapshot.edit();
        transaction.accept(0).expect("accept package revision");
        let commit = transaction.commit().expect("commit package revision");
        let published = package
            .apply_tracked_revisions(&commit)
            .expect("publish package revision");
        assert!(
            !std::str::from_utf8(published.source())
                .unwrap()
                .contains("w:ins")
        );

        let inverse = commit.patch().inverse();
        let restored = package
            .apply_tracked_revision_patch(&inverse)
            .expect("publish package inverse");
        assert_eq!(restored.source(), snapshot.source());
    }
}
