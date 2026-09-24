//! Immutable opened-package state and finite resource policy.

use std::collections::HashMap;
use std::fmt;
use std::sync::{Arc, OnceLock};

use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{OpcPackage, PackURI};
use sha2::{Digest, Sha256};

use crate::notes::SlideRootMemo;
use crate::parts::{MceCapture, PresentationPart, RetainedMce};
use crate::{Error, Result};

/// Maximum selector/value pairs in one atomic same-slide text batch.
pub const MAX_SHAPE_TEXT_REPLACEMENTS: usize = 256;

/// One borrowed selector/value pair in an atomic same-slide text batch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ShapeTextReplacement<'a> {
    selector: crate::shape::Key<'a>,
    text: &'a str,
}

impl<'a> ShapeTextReplacement<'a> {
    /// Build a replacement from an exact semantic selector.
    #[must_use]
    pub const fn new(selector: crate::shape::Key<'a>, text: &'a str) -> Self {
        Self { selector, text }
    }

    /// Select a shape by exact producer-visible name.
    #[must_use]
    pub const fn named(name: &'a str, text: &'a str) -> Self {
        Self::new(crate::shape::Key::Name(name), text)
    }

    /// Select a shape by checked zero-based pre-order position.
    #[must_use]
    pub const fn at(index: usize, text: &'a str) -> Self {
        Self::new(crate::shape::Key::Index(index), text)
    }

    /// Exact semantic selector for this replacement.
    #[must_use]
    pub const fn selector(self) -> crate::shape::Key<'a> {
        self.selector
    }

    /// Borrowed replacement text.
    #[must_use]
    pub const fn text(self) -> &'a str {
        self.text
    }
}

/// Finite limits for opened-presentation transactions and durable patches.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_parts: usize,
    max_patch_bytes: usize,
    max_text_bytes: usize,
    max_history_entries: usize,
    max_history_bytes: usize,
    max_retained_candidate_bytes: usize,
    max_retained_mce_bytes: usize,
}

impl Limits {
    /// Conservative defaults suitable for ordinary presentations.
    pub const DEFAULT: Self = Self {
        max_parts: 4_096,
        max_patch_bytes: 128 * 1024 * 1024,
        max_text_bytes: 8 * 1024 * 1024,
        max_history_entries: 64,
        max_history_bytes: 256 * 1024 * 1024,
        max_retained_candidate_bytes: 64 * 1024 * 1024,
        max_retained_mce_bytes: 1024 * 1024,
    };

    /// Construct a finite, nonzero policy.
    #[must_use]
    pub const fn new(
        max_parts: usize,
        max_patch_bytes: usize,
        max_text_bytes: usize,
        max_history_entries: usize,
        max_history_bytes: usize,
        max_retained_candidate_bytes: usize,
    ) -> Option<Self> {
        if max_parts == 0
            || max_patch_bytes == 0
            || max_text_bytes == 0
            || max_history_entries == 0
            || max_history_bytes == 0
            || max_retained_candidate_bytes == 0
        {
            None
        } else {
            Some(Self {
                max_parts,
                max_patch_bytes,
                max_text_bytes,
                max_history_entries,
                max_history_bytes,
                max_retained_candidate_bytes,
                max_retained_mce_bytes: Self::DEFAULT.max_retained_mce_bytes,
            })
        }
    }

    /// Maximum number of changed parts in one patch.
    #[must_use]
    pub const fn max_parts(self) -> usize {
        self.max_parts
    }

    /// Maximum encoded durable-patch size.
    #[must_use]
    pub const fn max_patch_bytes(self) -> usize {
        self.max_patch_bytes
    }

    /// Maximum replacement text size.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Maximum retained history entries.
    #[must_use]
    pub const fn max_history_entries(self) -> usize {
        self.max_history_entries
    }

    /// Maximum aggregate encoded history size.
    #[must_use]
    pub const fn max_history_bytes(self) -> usize {
        self.max_history_bytes
    }

    /// Maximum serialized candidate archive a cross-presentation slide-copy
    /// plan may retain for its own application.
    ///
    /// A plan whose candidate archive fits this ceiling keeps the archive it
    /// already built, and applying the plan reuses those bytes instead of
    /// serializing and deflating the candidate a second time.  A candidate
    /// above the ceiling is not retained: the plan is returned unretained and
    /// application rebuilds the archive exactly as it does without retention.
    /// Exceeding this ceiling is therefore never a refusal, because rebuilding
    /// is always available.
    ///
    /// The retained bytes are observable through
    /// [`CrossSlideCopyPlan::retained_candidate_bytes`] and released by
    /// [`CrossSlideCopyPlan::release_retained_candidate`] or by dropping the
    /// plan.  The default is 64 MiB.  This ceiling bounds what a plan *holds*.
    /// Applying the plan shares the retained archive with the package it
    /// opens and publishes rather than copying it (change 0751), so after an
    /// application the plan and the destination may hold one allocation
    /// between them, and releasing or dropping the plan frees it only once
    /// the destination no longer holds it.
    ///
    /// Zero is rejected for this field by [`Self::new`]. Pass `1` to turn
    /// candidate retention off, since no ZIP archive is
    /// one byte long.
    ///
    /// [`CrossSlideCopyPlan::retained_candidate_bytes`]: crate::opened::CrossSlideCopyPlan::retained_candidate_bytes
    /// [`CrossSlideCopyPlan::release_retained_candidate`]: crate::opened::CrossSlideCopyPlan::release_retained_candidate
    #[must_use]
    pub const fn max_retained_candidate_bytes(self) -> usize {
        self.max_retained_candidate_bytes
    }

    /// Set the optional retained slide-MCE projection budget in bytes.
    ///
    /// Zero disables retention. Exceeding this ceiling falls back to ordinary
    /// preprocessing and never changes a value, refusal, or published byte.
    /// Existing six-argument construction uses the one-MiB default.
    #[must_use]
    pub const fn with_max_retained_mce_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_mce_bytes = maximum;
        self
    }

    /// Maximum aggregate charge for retained slide-MCE projections.
    ///
    /// This separate one-MiB default ceiling covers owned output vector
    /// capacity and conservatively accounted memo metadata. It is not an RSS
    /// estimate. Zero disables retention; each snapshot reports its charge and
    /// can release its retained projections without changing its meaning.
    #[must_use]
    pub const fn max_retained_mce_bytes(self) -> usize {
        self.max_retained_mce_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Stable identity of one slide in current presentation order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Slide {
    pub(crate) id: u32,
    pub(crate) relationship_id: String,
    pub(crate) part_name: PackURI,
    pub(crate) name: String,
}

#[derive(Debug, Clone, Copy)]
struct SlideNameMatch {
    first: usize,
    count: usize,
}

/// Immutable exact-name lookup for one captured slide list.
///
/// Names are not required to be unique by the package model: the public
/// selector reports an ambiguity when more than one slide has the requested
/// name. Keeping the first position and match count preserves that behavior
/// without rescanning the slide list for every selector.
#[derive(Debug, Clone)]
pub(crate) struct SlideNameIndex {
    by_name: HashMap<String, SlideNameMatch>,
}

impl SlideNameIndex {
    pub(crate) fn build(slides: &[Slide]) -> Result<Self> {
        let mut by_name: HashMap<String, SlideNameMatch> = HashMap::new();
        by_name
            .try_reserve(slides.len())
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation slide name index",
                source,
            })?;
        for (index, slide) in slides.iter().enumerate() {
            if let Some(existing) = by_name.get_mut(slide.name.as_str()) {
                existing.count = existing.count.saturating_add(1);
                continue;
            }
            let mut name = String::new();
            name.try_reserve(slide.name.len())
                .map_err(|source| Error::Allocation {
                    resource: "opened-presentation slide name index key",
                    source,
                })?;
            name.push_str(&slide.name);
            by_name.insert(
                name,
                SlideNameMatch {
                    first: index,
                    count: 1,
                },
            );
        }
        Ok(Self { by_name })
    }

    pub(crate) fn resolve(&self, slides: &[Slide], name: &str) -> Result<Slide> {
        let Some(matched) = self.by_name.get(name) else {
            return Err(Error::SlideNameNotFound(name.to_owned()));
        };
        if matched.count != 1 {
            return Err(Error::AmbiguousSlideName {
                name: name.to_owned(),
                matches: matched.count,
            });
        }
        slides
            .get(matched.first)
            .cloned()
            .ok_or_else(|| invalid("opened-presentation slide name index lost its position"))
    }
}

impl Slide {
    /// Stable `p:sldId@id` identity.
    #[must_use]
    pub const fn id(&self) -> u32 {
        self.id
    }

    /// Exact producer-visible name, with the part name as fallback.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Physical slide part name retained across ordering edits.
    #[must_use]
    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }
}

/// Immutable root snapshot for an opened presentation.
#[derive(Clone)]
pub struct Snapshot {
    pub(crate) package: Arc<OpcPackage>,
    pub(crate) presentation_name: PackURI,
    pub(crate) slides: Vec<Slide>,
    pub(crate) slide_name_index: SlideNameIndex,
    pub(crate) revision: [u8; 32],
    pub(crate) limits: Limits,
    pub(crate) physical_source_provenance: bool,
    /// Serialized-archive revision of `package`, memoized with the archive
    /// bound it was taken under.
    ///
    /// The captured package is immutable behind an `Arc` and is never mutated
    /// after capture, so the serialized archive it publishes — and therefore
    /// its digest — is a pure function of this snapshot. Clones of a snapshot
    /// share the same package and the same cache; a rebind onto a different
    /// package starts an empty one, because content equality does not imply
    /// an identical retained archive.
    pub(crate) physical_revision: Arc<OnceLock<(usize, [u8; 32])>>,
    /// Per-part payload digests of `package`, keyed by payload allocation.
    ///
    /// Every entry names an allocation this snapshot's own `package` holds, so
    /// the memo pins no bytes the snapshot does not already own. It is an
    /// accelerator for [`package_fingerprint`] and nothing else: a miss is an
    /// ordinary hash and no value, refusal or published byte depends on a hit.
    pub(crate) part_digests: Arc<PartDigests>,
    /// Default-profile transformed slide XML retained for later captures.
    /// This table is optional, bounded by `Limits`, and shares cheaply across
    /// snapshot clones without changing the snapshot's semantic meaning.
    pub(crate) retained_mce: Option<Arc<RetainedMce>>,
    /// Notes-root classifications this snapshot's capture proved, keyed by
    /// the payload allocation each one read (ADR 0032).
    ///
    /// Every entry names an allocation `package` holds, so the memo pins no
    /// bytes the snapshot does not already own. A capture of a package that
    /// still shares a slide payload allocation — a commit's staged package
    /// shares every slide it did not rewrite — reuses the classification
    /// instead of rescanning that slide. A miss is the ordinary scan. The memo
    /// is built by the capture or projected by a rebind, never mutated, and
    /// clones share it.
    pub(crate) slide_roots: Arc<SlideRootMemo>,
}

/// Memoized per-part payload digests, keyed by payload allocation address and
/// length, retaining the payload `Arc` so the key cannot be recycled.
///
/// An entry asserts *the allocation at address `a` of length `l` has payload
/// digest `d`*. Retaining the `Arc` is load-bearing rather than an
/// optimization: without a strong reference the allocation could be freed and
/// a different payload allocated at the same address, and a lookup would then
/// answer with the digest of bytes that no longer exist. Holding the `Arc`
/// makes the address unrecyclable for the entry's lifetime, so a hit proves
/// allocation identity and therefore byte identity. The one degenerate case,
/// a zero-length payload, is safe in the other direction: two distinct empty
/// `Vec`s may share a dangling address and they have the same (empty) payload,
/// so such a hit still returns the right digest.
#[derive(Default)]
pub(crate) struct PartDigests {
    entries: HashMap<(usize, usize), (Arc<Vec<u8>>, [u8; 32])>,
}

impl PartDigests {
    fn with_capacity(parts: usize) -> Result<Self> {
        let mut entries = HashMap::new();
        entries
            .try_reserve(parts)
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation part digests",
                source,
            })?;
        Ok(Self { entries })
    }

    /// Digest memoized for exactly this allocation, if any.
    fn get(&self, key: (usize, usize)) -> Option<[u8; 32]> {
        self.entries.get(&key).map(|(_blob, digest)| *digest)
    }

    /// One parent memo answering for every allocation any of `parents`
    /// names, for a capture of a package assembled from several packages'
    /// payloads (change 0751).
    ///
    /// Each entry is copied with the payload `Arc` it retains, so every
    /// answer is still the digest of the allocation its own memo hashed. Two
    /// parents naming one allocation name the same bytes and hold the same
    /// digest. The union is transient: a capture consults it and keeps a memo
    /// built over its own package, so the union pins nothing beyond that
    /// capture. `None` when it cannot be reserved; a memo is optional, so the
    /// caller then captures without one.
    pub(crate) fn union(parents: &[&Self]) -> Option<Self> {
        let capacity = parents.iter().try_fold(0usize, |total, parent| {
            total.checked_add(parent.entries.len())
        })?;
        let mut union = Self::with_capacity(capacity).ok()?;
        for parent in parents {
            for (key, (blob, digest)) in &parent.entries {
                union.entries.insert(*key, (Arc::clone(blob), *digest));
            }
        }
        Some(union)
    }

    /// Return the package-owned payload allocation for an exact memo key.
    /// Retention callers use this after the owned package has been fingerprinted
    /// so foreign `blob_arc` implementations cannot donate an unverified owner.
    fn owner_for(&self, key: (usize, usize)) -> Option<Arc<Vec<u8>>> {
        self.entries
            .get(&key)
            .map(|(blob, _digest)| Arc::clone(blob))
    }

    /// [`Self::get`] for the ABA gate, which keys by address directly.
    #[cfg(test)]
    pub(crate) fn get_for_test(&self, key: (usize, usize)) -> Option<[u8; 32]> {
        self.get(key)
    }

    fn insert(&mut self, key: (usize, usize), blob: Arc<Vec<u8>>, digest: [u8; 32]) -> Result<()> {
        self.entries
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation part digests",
                source,
            })?;
        self.entries.insert(key, (blob, digest));
        Ok(())
    }

    /// Number of memoized payloads.
    #[cfg(test)]
    pub(crate) fn len(&self) -> usize {
        self.entries.len()
    }

    /// Memoized allocation keys, for the retention gate.
    #[cfg(test)]
    pub(crate) fn keys(&self) -> impl Iterator<Item = (usize, usize)> + '_ {
        self.entries.keys().copied()
    }

    /// Each entry's key with the payload allocation it retains, for the tests
    /// that prove an entry retains the very `Arc` a package holds.
    #[cfg(test)]
    pub(crate) fn retained(&self) -> impl Iterator<Item = ((usize, usize), &Arc<Vec<u8>>)> + '_ {
        self.entries
            .iter()
            .map(|(key, (blob, _digest))| (*key, blob))
    }

    /// Strong references the memo itself holds on each memoized payload.
    #[cfg(test)]
    pub(crate) fn strong_counts(&self) -> impl Iterator<Item = usize> + '_ {
        self.entries
            .values()
            .map(|(blob, _)| Arc::strong_count(blob))
    }

    /// Bytes this memo occupies itself, excluding the payloads it names.
    ///
    /// One `HashMap` slot is a key, a value and one control byte; the map
    /// grows in capacity steps, so the reported per-entry figure varies while
    /// the per-slot cost is flat.
    #[cfg(test)]
    pub(crate) fn resident_bytes(&self) -> usize {
        self.entries.capacity()
            * (size_of::<(usize, usize)>() + size_of::<(Arc<Vec<u8>>, [u8; 32])>() + 1)
    }

    /// Project `self` onto the blobs `package` currently holds.
    ///
    /// An entry survives only when `package` holds the very allocation it
    /// names, so the projection never pins a payload the package dropped.
    /// Each kept entry retains `package`'s own `Arc` for that allocation, never
    /// the one `self` retained, so the result is re-projected onto `package`'s
    /// allocations rather than inherited; the facade relies on this at every
    /// point it keeps a snapshot's memo (change 0751).
    pub(crate) fn project(&self, package: &OpcPackage) -> Result<Self> {
        if self.entries.is_empty() {
            return Ok(Self::default());
        }
        let mut projected = Self::with_capacity(package.part_count())?;
        for metadata in package.iter_parts() {
            let part = package.get_part(metadata.partname())?;
            let Some((key, blob)) = memo_key(part) else {
                continue;
            };
            if let Some(digest) = self.get(key) {
                projected.insert(key, blob, digest)?;
            }
        }
        Ok(projected)
    }
}

impl fmt::Debug for PartDigests {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PartDigests")
            .field("len", &self.entries.len())
            .finish()
    }
}

/// Memo key for `part`, or `None` when its `blob_arc()` does not alias the
/// bytes `blob()` returns.
///
/// `Part` is a public trait: a foreign implementation may return an `Arc`
/// whose contents differ from its visible payload, and `litchi-opc`'s
/// `from_vec_reusing_payloads` guards the same way before reusing a donor
/// payload. A part that fails this alias test is never memoized and is hashed
/// exactly as the unmemoized path hashes it, so an inconsistent part can cost
/// performance and can never change a value.
fn memo_key(part: &dyn litchi_opc::Part) -> Option<((usize, usize), Arc<Vec<u8>>)> {
    let blob = part.blob_arc();
    let visible = part.blob();
    if blob.len() != visible.len() || !std::ptr::eq(blob.as_slice(), visible) {
        return None;
    }
    let key = (blob.as_ptr() as usize, blob.len());
    Some((key, blob))
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("presentation_name", &self.presentation_name)
            .field("slides", &self.slides)
            .field("revision", &self.revision)
            .field("limits", &self.limits)
            .field("retained_mce_bytes", &self.retained_mce_bytes())
            .finish_non_exhaustive()
    }
}

impl Snapshot {
    /// Slides in current presentation order.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Fingerprint of the complete captured OPC graph.
    #[must_use]
    pub const fn revision(&self) -> [u8; 32] {
        self.revision
    }

    /// Start one detached atomic transaction.
    #[must_use]
    pub fn edit(&self) -> super::Transaction {
        super::Transaction::new(self.clone())
    }

    /// Rebind this snapshot's validated state onto a package that carries
    /// byte-identical [`package_fingerprint`] inputs.
    ///
    /// Every derived field — the presentation root, the slide identities, the
    /// name index and the complete-package revision — is a function of exactly
    /// the content [`packages_equal`] compares, so the result is the snapshot
    /// [`capture_with_provenance`] would return for `package`. Callers must
    /// establish that equality first.
    pub(crate) fn rebound_to(&self, package: &OpcPackage) -> Self {
        debug_assert!(
            packages_equal(self.package.as_ref(), package),
            "opened-presentation snapshot rebound to a different package"
        );
        let owned = Arc::new(package.clone());
        let part_digests = Arc::new(
            self.part_digests
                .project(owned.as_ref())
                .unwrap_or_else(|_error| PartDigests::default()),
        );
        let retained_mce = self
            .retained_mce
            .as_ref()
            .and_then(|retained| retained.project_with_owner(|key| part_digests.owner_for(key)));
        // Like the MCE table, the slide-root memo keeps only the entries whose
        // allocation the rebound package itself holds, each retaining that
        // package's own `Arc`; like the digest projection, a table that cannot
        // be reserved is empty, which only costs a later scan (ADR 0032).
        let slide_roots = Arc::new(self.slide_roots.project(|key| part_digests.owner_for(key)));
        Self {
            // `packages_equal` proves the fingerprint inputs are identical; it
            // says nothing about ZIP ordering, compression, or retained source
            // bytes, so the serialized-archive revision is not carried over.
            physical_revision: Arc::new(OnceLock::new()),
            // The part-digest memo claims only that a given payload
            // allocation hashes to a given digest, which `packages_equal` does
            // not disturb, so it is projected onto the rebound package's own
            // allocations rather than discarded. `project` drops every entry
            // the new package does not hold, so the memo still pins nothing
            // the snapshot does not own. A projection that cannot allocate
            // falls back to an empty memo, which only costs a later hash.
            part_digests,
            // The digest projection proves that each surviving MCE key still
            // names an allocation owned by the rebound package. Entries whose
            // source allocation was replaced are simply dropped.
            retained_mce,
            slide_roots,
            package: owned,
            ..self.clone()
        }
    }

    /// Aggregate charged bytes held by this snapshot's optional MCE table.
    ///
    /// Counts output vector capacity and table/Arc metadata. Clones may share
    /// these allocations, so charges across snapshots are not additive
    /// physical memory. This charge excludes the already-owned raw package.
    ///
    /// This is the charge for this owner. Snapshot clones may report the same
    /// value while sharing one physical table, so per-owner values are not
    /// additive physical-memory measurements.
    #[must_use]
    pub fn retained_mce_bytes(&self) -> usize {
        self.retained_mce
            .as_deref()
            .map_or(0, RetainedMce::retained_bytes)
    }

    /// Release this snapshot's reference to retained transformed slide XML
    /// while preserving every captured slide, revision, policy, and package
    /// byte. Snapshot clones may keep the shared table alive.
    pub fn release_retained_mce(&mut self) {
        self.retained_mce = None;
    }

    /// Resource policy inherited by edits and patches.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    pub(crate) fn resolve_slide(&self, key: crate::slide::Key<'_>) -> Result<Slide> {
        match key {
            crate::slide::Key::Index(index) => {
                self.slides
                    .get(index)
                    .cloned()
                    .ok_or(Error::SlideIndexOutOfBounds {
                        index,
                        len: self.slides.len(),
                    })
            },
            crate::slide::Key::Name(name) => self.slide_name_index.resolve(&self.slides, name),
        }
    }
}

pub(crate) fn capture(
    package: &OpcPackage,
    limits: Limits,
    physical_source_provenance: bool,
) -> Result<Snapshot> {
    capture_with_provenance(package, limits, physical_source_provenance)
}

pub(crate) fn capture_with_provenance(
    package: &OpcPackage,
    limits: Limits,
    physical_source_provenance: bool,
) -> Result<Snapshot> {
    capture_internal(
        package,
        limits,
        physical_source_provenance,
        Revision::Cold,
        None,
        None,
    )
}

/// Capture `package`, reusing `parent`'s payload digest for every payload
/// allocation the two packages share.
///
/// The revision is the one [`capture_with_provenance`] would compute: a memo
/// hit only declines to re-hash bytes whose allocation identity — and
/// therefore whose content — is already proven.
pub(crate) fn capture_with_parent_digests(
    package: &OpcPackage,
    limits: Limits,
    physical_source_provenance: bool,
    parent: &PartDigests,
) -> Result<Snapshot> {
    capture_internal(
        package,
        limits,
        physical_source_provenance,
        Revision::Parent(parent),
        None,
        None,
    )
}

/// Capture a package with its already computed revision, reusing a parent
/// snapshot's retained MCE projections and its proved notes-root
/// classifications for every payload allocation the two packages share. A
/// commit passes its source snapshot's; the cross-slide copy passes neither.
/// Every validation still runs; only work whose result is proved for the
/// identical bytes is not repeated.
pub(crate) fn capture_with_revision_and_digests_and_mce(
    package: &OpcPackage,
    limits: Limits,
    physical_source_provenance: bool,
    revision: [u8; 32],
    digests: PartDigests,
    parent_mce: Option<&RetainedMce>,
    parent_roots: Option<&SlideRootMemo>,
) -> Result<Snapshot> {
    capture_internal(
        package,
        limits,
        physical_source_provenance,
        Revision::Known(revision, digests),
        parent_mce,
        parent_roots,
    )
}

/// How one capture obtains its complete-package revision and its memo.
enum Revision<'a> {
    /// Compute it with no memo to consult.
    Cold,
    /// Compute it, reusing `parent`'s digest for every shared allocation.
    Parent(&'a PartDigests),
    /// Reuse a revision the caller computed from this exact content, with the
    /// memo it filled while computing it.
    Known([u8; 32], PartDigests),
}

fn capture_internal(
    package: &OpcPackage,
    limits: Limits,
    physical_source_provenance: bool,
    revision: Revision<'_>,
    parent_mce: Option<&RetainedMce>,
    parent_roots: Option<&SlideRootMemo>,
) -> Result<Snapshot> {
    let presentation = PresentationPart::from_package(package)?;
    let presentation_name = presentation.part().partname().clone();
    let references = presentation.slide_references()?;
    if references.len() > limits.max_parts {
        return Err(Error::Limit {
            resource: "opened-presentation slides",
            limit: limits.max_parts,
        });
    }
    let view = crate::presentation::Presentation::new(presentation, package);
    let mut mce_capture = MceCapture::new(parent_mce, limits.max_retained_mce_bytes());
    let captured = view.capture_slides_with_mce(&mut mce_capture, parent_roots)?;
    if references.len() != captured.slides.len() {
        return Err(invalid(
            "opened-presentation slide references do not resolve one-to-one",
        ));
    }
    let mut slides = Vec::new();
    slides
        .try_reserve_exact(references.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide identities",
            source,
        })?;
    let mut ids = std::collections::HashSet::new();
    let mut relationship_ids = std::collections::HashSet::new();
    let mut part_names = std::collections::HashSet::new();
    let mut first_name_error = captured.first_name_error;
    let slide_root_proofs = captured.slide_root_proofs;
    for (index, (reference, captured)) in references.iter().zip(captured.slides).enumerate() {
        let slide = captured.slide;
        let relationship = presentation
            .part()
            .rels()
            .get(reference.relationship_id())
            .ok_or_else(|| invalid("opened-presentation slide relationship is missing"))?;
        if relationship.is_external()
            || !crate::parts::is_relationship_type(relationship.reltype(), rt::SLIDE, "slide")
        {
            return Err(invalid(
                "opened-presentation slide relationship is unsupported",
            ));
        }
        let target = relationship.target_partname()?;
        let part_name = slide.part().part().partname().clone();
        if target != part_name {
            return Err(invalid(
                "opened-presentation slide relationship target changed during capture",
            ));
        }
        if !ids.insert(reference.id())
            || !relationship_ids.insert(reference.relationship_id().to_owned())
            || !part_names.insert(part_name.clone())
        {
            return Err(invalid(
                "opened-presentation slide identities are not one-to-one",
            ));
        }
        let name = if let Some(name) = captured.name {
            name
        } else {
            let Some((name_index, error)) = first_name_error.take() else {
                return Err(invalid(
                    "opened-presentation capture name projection is incomplete",
                ));
            };
            if name_index != index {
                return Err(invalid(
                    "opened-presentation capture name projection order changed",
                ));
            }
            return Err(error);
        };
        slides.push(Slide {
            id: reference.id(),
            relationship_id: reference.relationship_id().to_owned(),
            name,
            part_name,
        });
    }
    let _notes = match slide_root_proofs.as_deref() {
        Some(proofs) => crate::notes::load_snapshot_with_slide_root_proofs(
            package,
            &presentation_name,
            references.len(),
            proofs,
        )?,
        None => crate::notes::load_snapshot(package, &presentation_name)?,
    };
    let slide_name_index = SlideNameIndex::build(&slides)?;
    // The snapshot owns its package from here on, and the memo it keeps must
    // name that package's own payload allocations, so the revision is taken
    // over the owned clone rather than over the borrowed original.
    let owned = Arc::new(package.clone());
    let (revision, part_digests) = match revision {
        Revision::Known(revision, digests) => {
            debug_assert!(
                package_fingerprint(owned.as_ref()).is_ok_and(|fresh| fresh == revision),
                "opened-presentation capture reused a stale complete-package revision"
            );
            // The caller's memo was filled over the package it hashed. The
            // built-in parts share their payload `Arc` when a package is
            // cloned, so the projection keeps every entry; a foreign part that
            // copies its payload instead simply loses its entry, which costs a
            // later hash and can never answer wrongly.
            (
                revision,
                digests
                    .project(owned.as_ref())
                    .unwrap_or_else(|_error| PartDigests::default()),
            )
        },
        Revision::Cold => package_fingerprint_with_memo(owned.as_ref(), None)?,
        Revision::Parent(parent) => package_fingerprint_with_memo(owned.as_ref(), Some(parent))?,
    };
    let retained_mce = mce_capture.finish(|key| part_digests.owner_for(key));
    // The proofs borrow the input package. Each successful classification is
    // kept only when the owned package's digest memo holds, and hands over,
    // the very allocation it read; a refused reservation fails the capture
    // with a typed error after every validation has passed (ADR 0032).
    let slide_roots = Arc::new(SlideRootMemo::from_records(
        slide_root_proofs
            .as_deref()
            .unwrap_or_default()
            .iter()
            .map(|proof| proof.record()),
        |key| part_digests.owner_for(key),
    )?);
    drop(slide_root_proofs);
    Ok(Snapshot {
        package: owned,
        presentation_name,
        slides,
        slide_name_index,
        revision,
        limits,
        physical_source_provenance,
        physical_revision: Arc::new(OnceLock::new()),
        part_digests: Arc::new(part_digests),
        retained_mce,
        slide_roots,
    })
}

/// The `litchi-pptx-opened-v2` complete-package revision of `package`.
///
/// This is the complete-package proof every opened-presentation verdict is
/// taken over, and the value both durable patch families embed. It is a total
/// function of exactly the content [`packages_equal`] compares — the root
/// relationships, the opaque non-part members, and per part the name, the
/// content type, the payload and the relationships — and of nothing else: no
/// `Arc` identity, no insertion order, no ZIP layout and no compression.
pub(crate) fn package_fingerprint(package: &OpcPackage) -> Result<[u8; 32]> {
    Ok(package_fingerprint_with_memo(package, None)?.0)
}

/// The `litchi-pptx-opened-v2` complete-package revision, computed over sorted
/// per-part digests, together with the payload-digest memo it filled.
///
/// The proof is tiered. A payload digest covers exactly one part's payload
/// bytes; a part digest covers the part name, the content type, that payload
/// digest and the part's relationships; the revision covers the package header
/// and the sequence of part digests in sorted part-name order. Each tier
/// carries its own domain string, so a digest of one tier can never be read as
/// a digest of another.
///
/// Only the payload tier is memoized, and deliberately so: a part's content
/// type and relationships are not behind its payload `Arc`, so memoizing a
/// whole part digest on payload identity would return a stale digest for a
/// part whose relationships moved while its payload did not. Keying the memo
/// and its value over exactly the same bytes removes that trap; the name, the
/// content type and the relationships are re-fed on every pass.
///
/// A payload whose allocation `parent` already names is taken from the memo
/// instead of being hashed again. A miss is an ordinary hash, so the returned
/// revision does not depend on which entries `parent` happened to hold.
pub(crate) fn package_fingerprint_with_memo(
    package: &OpcPackage,
    parent: Option<&PartDigests>,
) -> Result<([u8; 32], PartDigests)> {
    let mut parts = Vec::new();
    parts
        .try_reserve_exact(package.part_count())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation fingerprint parts",
            source,
        })?;
    parts.extend(package.iter_parts());
    parts.sort_unstable_by(|left, right| left.partname().as_str().cmp(right.partname().as_str()));
    let mut digest = Sha256::new();
    feed(&mut digest, b"litchi-pptx-opened-v2");
    let mut root_relationships = Vec::new();
    root_relationships
        .try_reserve_exact(package.rels().len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation fingerprint root relationships",
            source,
        })?;
    root_relationships.extend(package.rels().iter());
    root_relationships.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    feed_relationships(&mut digest, &root_relationships)?;
    let non_part_count = u32::try_from(package.non_part_members().len())
        .map_err(|_error| invalid("opened-presentation non-part member count exceeds u32"))?;
    feed(&mut digest, b"non-part-members");
    feed(&mut digest, &non_part_count.to_le_bytes());
    for member in package.non_part_members() {
        feed(&mut digest, member.name().as_bytes());
        feed(&mut digest, member.reason().as_str().as_bytes());
    }
    let part_count = u32::try_from(parts.len())
        .map_err(|_error| invalid("opened-presentation part count exceeds u32"))?;
    feed(&mut digest, b"parts");
    digest.update(part_count.to_le_bytes());
    let mut memo = PartDigests::with_capacity(parts.len())?;
    #[cfg(test)]
    let (part_total, mut hashed) = (parts.len(), 0usize);
    for metadata in parts {
        let part = package.get_part(metadata.partname())?;
        let payload = match memo_key(part) {
            Some((key, blob)) => {
                let payload = match parent.and_then(|parent| parent.get(key)) {
                    Some(memoized) => {
                        // Test and debug builds re-derive every reused digest,
                        // so the crate's own suite proves value identity
                        // rather than only that the memo compiles.
                        debug_assert_eq!(
                            memoized,
                            payload_digest(&blob),
                            "opened-presentation part-digest memo answered for different bytes"
                        );
                        memoized
                    },
                    None => {
                        #[cfg(test)]
                        {
                            hashed += 1;
                        }
                        payload_digest(&blob)
                    },
                };
                memo.insert(key, blob, payload)?;
                payload
            },
            None => {
                #[cfg(test)]
                {
                    hashed += 1;
                }
                payload_digest(part.blob())
            },
        };
        digest.update(part_digest(part, payload)?);
    }
    #[cfg(test)]
    fingerprint_log::record(parent.is_some(), part_total, hashed);
    Ok((digest.finalize().into(), memo))
}

/// Digest of one part payload, under its own domain string.
fn payload_digest(blob: &[u8]) -> [u8; 32] {
    let mut digest = Sha256::new();
    feed(&mut digest, b"litchi-pptx-opened-payload-v2");
    feed(&mut digest, blob);
    digest.finalize().into()
}

/// Digest of one part: its name, content type, payload digest and
/// relationships, under its own domain string.
fn part_digest(part: &dyn litchi_opc::Part, payload: [u8; 32]) -> Result<[u8; 32]> {
    let mut digest = Sha256::new();
    feed(&mut digest, b"litchi-pptx-opened-part-v2");
    feed(&mut digest, part.partname().as_str().as_bytes());
    feed(&mut digest, part.content_type().as_bytes());
    digest.update(payload);
    let mut relationships = Vec::new();
    relationships
        .try_reserve_exact(part.rels().len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation fingerprint relationships",
            source,
        })?;
    relationships.extend(part.rels().iter());
    relationships.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    feed_relationships(&mut digest, &relationships)?;
    Ok(digest.finalize().into())
}

/// The superseded `litchi-pptx-opened-v1` revision, retained only for the
/// differential gate that proves v1 and v2 return the same pairwise verdicts.
///
/// No production path computes this value: the durable formats that embedded
/// it carry superseded magics and are refused at parse.
#[cfg(test)]
pub(crate) fn package_fingerprint_v1(package: &OpcPackage) -> Result<[u8; 32]> {
    let mut parts = Vec::new();
    parts
        .try_reserve_exact(package.part_count())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation fingerprint parts",
            source,
        })?;
    // The fingerprint feeds every payload, so every payload is decoded here
    // (ADR 0030).
    for part in package.try_iter_parts() {
        parts.push(part?);
    }
    parts.sort_unstable_by(|left, right| left.partname().as_str().cmp(right.partname().as_str()));
    let mut digest = Sha256::new();
    feed(&mut digest, b"litchi-pptx-opened-v1");
    let mut root_relationships = Vec::new();
    root_relationships
        .try_reserve_exact(package.rels().len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation fingerprint root relationships",
            source,
        })?;
    root_relationships.extend(package.rels().iter());
    root_relationships.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    feed_relationships(&mut digest, &root_relationships)?;
    let non_part_count = u32::try_from(package.non_part_members().len())
        .map_err(|_error| invalid("opened-presentation non-part member count exceeds u32"))?;
    feed(&mut digest, b"non-part-members");
    feed(&mut digest, &non_part_count.to_le_bytes());
    for member in package.non_part_members() {
        feed(&mut digest, member.name().as_bytes());
        feed(&mut digest, member.reason().as_str().as_bytes());
    }
    for part in parts {
        feed(&mut digest, part.partname().as_str().as_bytes());
        feed(&mut digest, part.content_type().as_bytes());
        feed(&mut digest, part.blob());
        let mut relationships = Vec::new();
        relationships
            .try_reserve_exact(part.rels().len())
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation fingerprint relationships",
                source,
            })?;
        relationships.extend(part.rels().iter());
        relationships.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
        feed_relationships(&mut digest, &relationships)?;
    }
    Ok(digest.finalize().into())
}

/// Whether two packages present byte-identical [`package_fingerprint`] inputs.
///
/// Every input the fingerprint feeds is compared here in the same scope: the
/// package-root relationships, the opaque non-part members, and, per part, the
/// part name, the content type, the payload and the part relationships. A
/// `true` result therefore proves the two packages carry the same
/// complete-package revision without hashing either of them, and proves it by
/// content rather than by digest. A `false` result proves nothing at all, so
/// every caller falls back to the ordinary capture.
///
/// Payload comparison short-circuits on `Arc` pointer identity, which the
/// opened-transaction path preserves for every part an edit did not rewrite.
pub(crate) fn packages_equal(left: &OpcPackage, right: &OpcPackage) -> bool {
    if left.part_count() != right.part_count()
        || !relationships_equal(left.rels(), right.rels())
        || left.non_part_members().len() != right.non_part_members().len()
    {
        return false;
    }
    if left
        .non_part_members()
        .iter()
        .zip(right.non_part_members())
        .any(|(left, right)| left.name() != right.name() || left.reason() != right.reason())
    {
        return false;
    }
    // A payload that cannot be decoded proves nothing, so it reads as unequal
    // and the caller falls back to the ordinary capture (ADR 0030).
    left.try_iter_parts().all(|part| {
        part.is_ok_and(|part| {
            right.get_part(part.partname()).is_ok_and(|other| {
                part.content_type() == other.content_type()
                    && blobs_equal(part, other)
                    && relationships_equal(part.rels(), other.rels())
            })
        })
    })
}

fn blobs_equal(left: &dyn litchi_opc::Part, right: &dyn litchi_opc::Part) -> bool {
    Arc::ptr_eq(&left.blob_arc(), &right.blob_arc()) || left.blob() == right.blob()
}

fn relationships_equal(
    left: &litchi_opc::Relationships,
    right: &litchi_opc::Relationships,
) -> bool {
    left.len() == right.len()
        && left.iter().all(|relationship| {
            right.get(relationship.r_id()).is_some_and(|other| {
                other.reltype() == relationship.reltype()
                    && other.target_ref() == relationship.target_ref()
                    && other.is_external() == relationship.is_external()
            })
        })
}

fn feed_relationships(
    digest: &mut Sha256,
    relationships: &[&litchi_opc::Relationship],
) -> Result<()> {
    let count = u32::try_from(relationships.len())
        .map_err(|_err| invalid("opened-presentation relationship count exceeds u32"))?;
    digest.update(count.to_le_bytes());
    for relationship in relationships {
        feed(digest, relationship.r_id().as_bytes());
        feed(digest, relationship.reltype().as_bytes());
        feed(digest, relationship.target_ref().as_bytes());
        digest.update([u8::from(relationship.is_external())]);
    }
    Ok(())
}

fn feed(digest: &mut Sha256, value: &[u8]) {
    digest.update(u64::try_from(value.len()).unwrap_or(u64::MAX).to_le_bytes());
    digest.update(value);
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

/// Which complete-package fingerprints a test thread computed, and how many
/// payloads each hashed rather than answering from a memo (change 0751).
///
/// Only fingerprints outside a `debug_assert!` are of interest to the tests
/// that read this log, and those are the ones taken with a parent memo; the
/// memo-free fingerprints that debug builds use to re-derive a value are
/// logged too and are told apart by `with_parent`.
#[cfg(test)]
pub(crate) mod fingerprint_log {
    use std::cell::RefCell;

    /// One fingerprint: whether a parent memo was consulted, the package's
    /// part count, and the payloads hashed rather than answered.
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub(crate) struct Entry {
        pub(crate) with_parent: bool,
        pub(crate) parts: usize,
        pub(crate) hashed: usize,
    }

    thread_local! {
        static LOG: RefCell<Vec<Entry>> = const { RefCell::new(Vec::new()) };
    }

    pub(super) fn record(with_parent: bool, parts: usize, hashed: usize) {
        LOG.with(|log| {
            log.borrow_mut().push(Entry {
                with_parent,
                parts,
                hashed,
            });
        });
    }

    /// Take this thread's log, leaving it empty.
    pub(crate) fn take() -> Vec<Entry> {
        LOG.with(|log| std::mem::take(&mut *log.borrow_mut()))
    }

    /// Take this thread's log and keep the fingerprints taken with a parent.
    pub(crate) fn take_with_parent() -> Vec<Entry> {
        take()
            .into_iter()
            .filter(|entry| entry.with_parent)
            .collect()
    }
}

#[cfg(test)]
pub(crate) mod mce_retention_test_hooks {
    use super::*;

    pub(crate) fn slide_output(snapshot: &Snapshot, index: usize) -> Option<Arc<Vec<u8>>> {
        let slide = snapshot.slides.get(index)?;
        let part = snapshot.package.get_part(&slide.part_name).ok()?;
        snapshot.retained_mce.as_deref()?.lookup(part.blob())
    }

    pub(crate) fn checked_charge_for_test(
        entries_capacity: usize,
        output_capacity: usize,
    ) -> Option<usize> {
        crate::parts::checked_mce_charge_for_test(entries_capacity, output_capacity)
    }
}
