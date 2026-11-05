//! Shared plumbing of the scoped PPT edit families (record 0357).
//!
//! Every scoped family `x` follows the record 0351 §2.3 shape:
//! `Package::edit_x(&self, ..) -> x::Edit`, `x::Edit::commit(self) ->
//! x::Commit` publishing a new [`Package`], and `Package::apply_x(&self,
//! &x::Patch) -> x::Commit`. Dropping an `Edit` discards it. The families
//! keep their record-level codecs and staging owners crate-private; this
//! module retains the source package, lazily captures exact bytes when a
//! family first needs them, publishes candidate bytes as a new package, and
//! replays byte patches.
//!
//! A family's patch is exact and record-sized (record 0357a): it carries the
//! length and block root (record 0357b, [`digest`]) of the source and target
//! package and the coalesced 512-byte blocks that differ, so `apply_x`
//! refuses a package that is not the patch's exact source
//! ([`Failure::PatchConflict`]), rebuilds the exact target, and `inverse`
//! swaps the endpoints. An exact no-op commit
//! publishes the source package unchanged (signatures included); a changed
//! commit of a signed package is refused ([`Failure::Signed`]).

#[cfg(test)]
use std::io::Cursor;
use std::sync::{Arc, Mutex, OnceLock};

use litchi_core::SourceVersion;
use litchi_core::edit::{Diagnostics, Refusal, RefusalKind};

use crate::package::{Error as PackageError, OpenOptions, Package};
use crate::slide_order;

mod digest;

pub(crate) use digest::ExactCache;
use digest::{BlockDigests, BlockSpan, Digest, Recipe};

/// A fallible, retryable lazy value.
///
/// The source package is immutable, but opening one of its native projections
/// can still fail.  Keeping the failure outside the cell means a failed
/// attempt returns its original typed error and does not require a panic or a
/// dummy value merely to satisfy `OnceLock`'s infallible initializer.  The
/// mutex is only held while the first successful value is built; the common
/// initialized path is one `OnceLock` read.
pub(crate) struct Lazy<T> {
    value: OnceLock<T>,
    initializing: Mutex<()>,
}

impl<T> Lazy<T> {
    pub(crate) const fn new() -> Self {
        Self {
            value: OnceLock::new(),
            initializing: Mutex::new(()),
        }
    }

    pub(crate) fn get_or_try_init<E>(
        &self,
        mut initialize: impl FnMut() -> Result<T, E>,
    ) -> Result<&T, E> {
        if let Some(value) = self.value.get() {
            return Ok(value);
        }
        let _guard = match self.initializing.lock() {
            Ok(guard) => guard,
            Err(poisoned) => poisoned.into_inner(),
        };
        loop {
            if let Some(value) = self.value.get() {
                return Ok(value);
            }
            let value = initialize()?;
            if self.value.set(value).is_ok() {
                // The successful set is guarded above, so the value is
                // necessarily visible here.  Keep the branch explicit rather
                // than relying on an unchecked unwrap in a public operation.
                if let Some(value) = self.value.get() {
                    return Ok(value);
                }
            }
        }
    }

    pub(crate) fn get_or_try_init_mut<E>(
        &mut self,
        mut initialize: impl FnMut() -> Result<T, E>,
    ) -> Result<&mut T, E> {
        if self.value.get().is_none() {
            let value = initialize()?;
            let _ = self.value.set(value);
        }
        // `&mut self` excludes a concurrent initializer, so the successful
        // fallible initialization above leaves the cell populated.  Keep the
        // invariant local rather than retrying through a borrow-checker loop
        // that could retain a mutable borrow across another initialization.
        Ok(self
            .value
            .get_mut()
            .expect("fallible lazy initialization populated its cell"))
    }

    pub(crate) fn get(&self) -> Option<&T> {
        self.value.get()
    }

    pub(crate) fn into_inner(self) -> Option<T> {
        self.value.into_inner()
    }
}

impl<T> std::fmt::Debug for Lazy<T> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Lazy")
            .field("initialized", &self.value.get().is_some())
            .finish()
    }
}

/// One side of an exact byte patch: the package length and its block root
/// (record 0357b).
#[derive(Clone, PartialEq, Eq)]
struct Endpoint {
    len: usize,
    digest: Digest,
}

impl Endpoint {
    fn new(len: usize, digest: Digest) -> Self {
        Self { len, digest }
    }
}

/// One coalesced changed range: the bytes at `offset` before and after the
/// publication that made the patch.
///
/// `before` and `after` have the same length except in the last range of a
/// patch whose endpoints differ in length, which runs to the end of each
/// endpoint.
#[derive(Clone, PartialEq, Eq)]
struct Range {
    offset: usize,
    before: Box<[u8]>,
    after: Box<[u8]>,
}

/// The comparison granularity of [`BytePatch::ranges`]: one OLE sector.
const BLOCK: usize = 512;

/// An exact, record-sized patch between two package byte images.
///
/// It carries both endpoints' length and block root and the coalesced 512-byte
/// blocks that differ, so it applies only to the exact source (anything else
/// is [`Failure::PatchConflict`]), reproduces the exact target, and inverts
/// by swapping the endpoints and reading each range's other side. The ranges
/// are shared between a patch, its clones and its inverses.
#[derive(Clone)]
pub(crate) struct BytePatch {
    source: Endpoint,
    target: Endpoint,
    ranges: Arc<[Range]>,
    block_spans: Arc<[BlockSpan]>,
    /// Whether the patch runs from each range's `after` to its `before`.
    inverted: bool,
}

impl PartialEq for BytePatch {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source
            && self.target == other.target
            && self.ranges.len() == other.ranges.len()
            && self
                .ranges
                .iter()
                .zip(other.ranges.iter())
                .all(|(left, right)| {
                    left.offset == right.offset && self.sides(left) == other.sides(right)
                })
    }
}

impl Eq for BytePatch {}

impl BytePatch {
    /// The exact no-op patch of a package of `len` bytes with root `digest`.
    fn noop(len: usize, digest: Digest) -> Self {
        let endpoint = Endpoint::new(len, digest);
        Self {
            source: endpoint.clone(),
            target: endpoint,
            ranges: Arc::from([]),
            block_spans: Arc::from([]),
            inverted: false,
        }
    }

    /// The changed ranges from `before` to `after`: compare in 512-byte
    /// blocks and coalesce adjacent changed blocks. Bytes past the shorter
    /// image belong to the last range. Equal images have no ranges.
    fn ranges(before: &[u8], after: &[u8]) -> Arc<[Range]> {
        let end = before.len().max(after.len());
        let side = |bytes: &[u8], start: usize, stop: usize| -> Box<[u8]> {
            bytes
                .get(start.min(bytes.len())..stop.min(bytes.len()))
                .unwrap_or_default()
                .into()
        };
        let range = |start: usize, stop: usize| Range {
            offset: start,
            before: side(before, start, stop),
            after: side(after, start, stop),
        };
        let mut ranges = Vec::new();
        let mut open: Option<usize> = None;
        let mut offset = 0;
        while offset < end {
            let stop = offset.saturating_add(BLOCK).min(end);
            let same = before
                .get(offset..stop)
                .is_some_and(|block| after.get(offset..stop) == Some(block));
            match (same, open) {
                (false, None) => open = Some(offset),
                (true, Some(start)) => {
                    ranges.push(range(start, offset));
                    open = None;
                },
                _ => {},
            }
            offset = stop;
        }
        if let Some(start) = open {
            ranges.push(range(start, end));
        }
        Arc::from(ranges)
    }

    /// The source and target sides of `range` in this patch's direction.
    fn sides<'range>(&self, range: &'range Range) -> (&'range [u8], &'range [u8]) {
        if self.inverted {
            (&range.after, &range.before)
        } else {
            (&range.before, &range.after)
        }
    }

    pub(crate) fn is_noop(&self) -> bool {
        self.ranges.is_empty()
    }

    pub(crate) fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            ranges: Arc::clone(&self.ranges),
            block_spans: Arc::clone(&self.block_spans),
            inverted: !self.inverted,
        }
    }

    /// The length of the package the patch applies to.
    pub(crate) fn source_len(&self) -> usize {
        self.source.len
    }

    /// The length of the package the patch produces.
    pub(crate) fn target_len(&self) -> usize {
        self.target.len
    }

    /// The block-index spans the changed ranges cover, sorted and disjoint:
    /// the target differs from the source only in these blocks.
    fn spans(&self) -> &[BlockSpan] {
        &self.block_spans
    }

    /// The bytes the changed ranges span (the longer side of each range).
    pub(crate) fn changed_bytes(&self) -> usize {
        self.ranges
            .iter()
            .map(|range| range.before.len().max(range.after.len()))
            .sum()
    }

    /// The bytes the patch retains: both sides of every range.
    fn payload_bytes(&self) -> usize {
        self.ranges
            .iter()
            .map(|range| range.before.len() + range.after.len())
            .sum()
    }

    /// Whether a package of `len` bytes with block root `digest` is the
    /// patch's exact source.
    fn is_source(&self, len: usize, digest: &Digest) -> bool {
        self.source.len == len && &self.source.digest == digest
    }

    /// The exact target rebuilt from the exact source bytes: resize the owned
    /// source buffer to the target length and write every range's target side.
    /// The caller has checked [`Self::is_source`]; with the source bound by
    /// length and block root the result is the exact target by construction
    /// (debug builds re-hash it).
    fn rebuild(&self, mut target: Vec<u8>) -> Option<Vec<u8>> {
        target.resize(self.target.len, 0);
        for range in self.ranges.iter() {
            let (_, after) = self.sides(range);
            let stop = range.offset.checked_add(after.len())?;
            target.get_mut(range.offset..stop)?.copy_from_slice(after);
        }
        debug_assert!(BlockDigests::of(&target).root() == self.target.digest);
        Some(target)
    }
}

/// The block-index spans of `ranges` (block-aligned offsets; a range covers
/// the blocks its longer side touches).
fn spans(ranges: &[Range]) -> Vec<BlockSpan> {
    ranges
        .iter()
        .map(|range| {
            let stop = range
                .offset
                .saturating_add(range.before.len().max(range.after.len()));
            BlockSpan::new(range.offset / BLOCK, stop.div_ceil(BLOCK))
        })
        .collect()
}

impl std::fmt::Debug for BytePatch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("source_len", &self.source.len)
            .field("target_len", &self.target.len)
            .field("ranges", &self.ranges.len())
            .field("changed_bytes", &self.changed_bytes())
            .field("payload_bytes", &self.payload_bytes())
            .field("block_span_count", &self.block_spans.len())
            .field(
                "block_span_bytes",
                &(self
                    .block_spans
                    .len()
                    .saturating_mul(size_of::<BlockSpan>())),
            )
            .finish()
    }
}

/// A scoped-family failure before it is wrapped in the family's own `Error`.
#[derive(Debug)]
pub(crate) enum Failure {
    Package(PackageError),
    Refused(Refusal),
    PatchConflict,
    Signed,
}

impl From<PackageError> for Failure {
    fn from(error: PackageError) -> Self {
        Self::Package(error)
    }
}

impl From<slide_order::Error> for Failure {
    fn from(error: slide_order::Error) -> Self {
        match error {
            slide_order::Error::Package(error) => Self::Package(error),
            slide_order::Error::Refused(slide_order::Refusal::Signed) => Self::Signed,
            slide_order::Error::Refused(refusal) => Self::Refused(Refusal::new(
                refusal.kind(),
                "publish",
                "the presentation cannot publish this scoped edit",
            )),
            slide_order::Error::Text(error) => match *error {
                crate::text_edit::Error::Package(error) => Self::Package(error),
                crate::text_edit::Error::Refused(refusal) => Self::Refused(Refusal::new(
                    refusal.kind(),
                    "publish",
                    "the presentation text cannot publish this scoped edit",
                )),
                crate::text_edit::Error::Source(_) => Self::Package(PackageError::Corrupted(
                    "scoped PPT publication failed in the source overlay".into(),
                )),
                crate::text_edit::Error::PatchConflict => Self::PatchConflict,
                crate::text_edit::Error::Signed => Self::Signed,
            },
            slide_order::Error::Composition(_) => Self::Package(PackageError::Corrupted(
                "scoped PPT publication failed to compose".into(),
            )),
        }
    }
}

/// Map a staging failure of `operation`: a value the record grammar refuses
/// becomes a content-free [`RefusalKind::InvalidValue`] refusal; limits and
/// I/O keep their package error.
pub(crate) fn staging(operation: &'static str) -> impl FnOnce(PackageError) -> Failure {
    move |error| match error {
        PackageError::InvalidFormat(_) => Failure::Refused(Refusal::new(
            RefusalKind::InvalidValue,
            operation,
            "the record grammar cannot hold the requested value",
        )),
        error => Failure::Package(error),
    }
}

/// Map a publication failure: content the record publisher cannot preserve
/// becomes a content-free [`RefusalKind::UnsupportedContent`] refusal.
pub(crate) fn publication(operation: &'static str) -> impl FnOnce(PackageError) -> Failure {
    move |error| match error {
        PackageError::InvalidFormat(_) => Failure::Refused(Refusal::new(
            RefusalKind::UnsupportedContent,
            operation,
            "the package holds content this publication cannot preserve",
        )),
        error => Failure::Package(error),
    }
}

/// A refusal of a target that does not exist.
pub(crate) fn not_found(operation: &'static str, reason: &'static str) -> Failure {
    Failure::Refused(Refusal::new(RefusalKind::NotFound, operation, reason))
}

/// A refusal of a target that cannot carry the edit.
pub(crate) fn unsupported_target(operation: &'static str, reason: &'static str) -> Failure {
    Failure::Refused(Refusal::new(
        RefusalKind::UnsupportedTarget,
        operation,
        reason,
    ))
}

/// One published scoped commit before it is wrapped in the family's types.
pub(crate) struct Published {
    pub(crate) package: Package,
    pub(crate) patch: BytePatch,
    pub(crate) diagnostics: Diagnostics,
}

/// The immutable source package a scoped edit started from, with a retryable
/// lazy handle for its exact bytes.
#[derive(Clone)]
pub(crate) struct Base {
    package: Package,
    expected_source: SourceVersion,
    bytes: Arc<Lazy<Arc<[u8]>>>,
}

impl std::fmt::Debug for Base {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Base")
            .field("package", &self.package)
            .finish_non_exhaustive()
    }
}

impl Base {
    /// Capture `package` for a scoped edit. Encrypted packages are refused:
    /// their content cannot be republished without storing it in plaintext.
    pub(crate) fn capture(package: &Package) -> Result<Self, Failure> {
        if package.is_encrypted() {
            return Err(Failure::Package(PackageError::EncryptionPolicy));
        }
        // Keep the source fence and encryption policy at entry, but defer the
        // exact-byte copy until a native editor, digest or publication really
        // needs it.  Cloning the package and the lazy handle is O(1).
        let expected_source = package.source_version()?;
        Ok(Self {
            package: package.clone(),
            expected_source,
            bytes: Arc::new(Lazy::new()),
        })
    }

    /// Borrow the retained immutable package for a projection that already
    /// lives in the package owner.  This accessor does not materialize the
    /// exact source bytes.
    pub(crate) const fn package(&self) -> &Package {
        &self.package
    }

    /// Reconstruct the exact low-level CFB profile retained by the package's
    /// shared source owner for source-backed editor ingress.
    pub(crate) fn cfb_limits(&self) -> Result<litchi_cfb::OleFileLimits, Failure> {
        let shared = self.package.shared_ole_file().limits();
        litchi_cfb::OleFileLimits::new(shared.max_input_bytes())
            .and_then(|limits| limits.with_max_directory_bytes(shared.max_directory_bytes()))
            .and_then(|limits| {
                limits.with_max_allocation_table_bytes(shared.max_allocation_table_bytes())
            })
            .map_err(|error| Failure::Package(PackageError::Ole(error)))
    }

    /// The exact source package bytes.
    pub(crate) fn bytes(&self) -> Result<&[u8], Failure> {
        Ok(self.bytes_arc()?.as_ref())
    }

    fn bytes_arc(&self) -> Result<&Arc<[u8]>, Failure> {
        self.ensure_source_version()?;
        let bytes = self.bytes.get_or_try_init(|| {
            self.ensure_source_version()?;
            let bytes = self.package.to_encrypted_bytes().map_err(Failure::from)?;
            self.ensure_source_version()?;
            Ok::<Arc<[u8]>, Failure>(Arc::from(bytes))
        })?;
        self.ensure_source_version()?;
        Ok(bytes)
    }

    fn ensure_source_version(&self) -> Result<(), Failure> {
        let observed = self.package.source_version()?;
        if observed != self.expected_source {
            return Err(Failure::Package(PackageError::Ole(
                litchi_cfb::OleError::SourceChanged {
                    expected: self.expected_source,
                    observed,
                },
            )));
        }
        Ok(())
    }

    /// The block root of the exact source bytes, cached on the package.
    fn digest(&self) -> Result<Digest, Failure> {
        Ok(self.package.exact().root(self.bytes()?))
    }

    /// Whether the source carries a digital-signature component, cached on
    /// the package.
    pub(crate) fn is_signed(&self) -> Result<bool, Failure> {
        self.ensure_source_version()?;
        source_signed(&self.package)
    }

    /// A shared handle on the exact source package bytes.
    pub(crate) fn shared_bytes(&self) -> Result<Arc<[u8]>, Failure> {
        Ok(Arc::clone(self.bytes_arc()?))
    }

    /// The options every republished package is opened with: the source
    /// package's limits and leniency.
    pub(crate) fn open_options(&self) -> OpenOptions<'static> {
        open_options(&self.package)
    }

    /// The source as the native slide-order read view, for families whose
    /// records live in the live document or a slide.
    pub(crate) fn native(&self) -> Result<slide_order::Snapshot, Failure> {
        let bytes = self.shared_bytes()?;
        let snapshot = slide_order::Snapshot::from_package_with_options(
            &self.package,
            bytes,
            self.open_options(),
        );
        // Preserve the edit-entry source fence even when the native snapshot
        // reused cached package projections whose own check observed a newer
        // baseline during construction.
        self.ensure_source_version()?;
        Ok(snapshot?)
    }

    /// The exact current bytes of one persisted owner record.
    pub(crate) fn owner_record(
        &self,
        native: &slide_order::Snapshot,
        owner: slide_order::records::Owner,
    ) -> Result<Vec<u8>, Failure> {
        Ok(slide_order::records::owner_record(native, owner)?)
    }

    /// Publish whole candidate package bytes.
    ///
    /// The patch is the coalesced byte-range diff against the source, bound
    /// to both endpoints' length and block root. The target's block list is
    /// derived from the source's (cached on the source package) by hashing
    /// only the changed and appended blocks, and seeds the published
    /// package's cache, so a later edit or patch of it hashes nothing it
    /// did not change.
    ///
    /// This is an internal, closed publication boundary. Its callers are the
    /// typed `font`, `picture_bullets`, `document_comparison`, `animation`
    /// and `embedded::object` editors, plus [`Self::publish_owner`]. The
    /// former editors emit a validated record projection; the latter first
    /// uses `slide_order::records::publish`, which rewrites one admitted
    /// owner record and verifies the reopened result. The embedded-object
    /// edit exposes only the typed `ExObjList` and generated persisted-object
    /// stream operations; its lower editor's unrelated stream helper is not
    /// reachable from this publication path. Every changed source is
    /// authenticated by
    /// [`Self::is_signed`] before the target is opened, and each editor
    /// refuses a changed signed source.
    ///
    /// `Base`, `BytePatch` and the family patch payloads are crate-private;
    /// the public family patches have a private field and no byte decoder.
    /// Thus a downstream caller cannot inject arbitrary package bytes into
    /// this boundary or make replay manufacture a new target. If a future
    /// publisher can import a package or add a reachable signature entry, it
    /// must not use the unsigned-cache seed without an equivalent proof.
    pub(crate) fn publish(self, after: Vec<u8>) -> Result<Published, Failure> {
        self.admit_output_length(after.len())?;
        let bytes = self.shared_bytes()?;
        let ranges = BytePatch::ranges(&bytes, &after);
        if ranges.is_empty() {
            return self.unchanged();
        }
        let block_spans = Arc::from(spans(&ranges));
        let source = self.package.exact().blocks(&bytes);
        let target = Arc::new(source.derive(&after, &block_spans));
        let patch = BytePatch {
            source: Endpoint::new(bytes.len(), source.root()),
            target: Endpoint::new(after.len(), target.root()),
            ranges,
            block_spans,
            inverted: false,
        };
        self.publish_patched(after, patch, Seed::Blocks(target))
    }

    /// Admit the candidate before constructing compact spans or block lists.
    /// The source is already admitted, so both endpoints then fit the compact
    /// index bound. Preserve the publication path's freshness/signature error
    /// precedence for an oversized changed candidate.
    fn admit_output_length(&self, observed: usize) -> Result<(), Failure> {
        self.ensure_source_version()?;
        admit_output_length(&self.package, observed)
    }

    /// Publish the exact target `after` of `patch`, made from this source,
    /// and seed the published package's digest cache.
    fn publish_patched(
        self,
        after: Vec<u8>,
        patch: BytePatch,
        seed: Seed,
    ) -> Result<Published, Failure> {
        publish_patched(&self.package, after, patch, seed)
    }

    /// Publish one persisted owner record replacement through the
    /// append-only native publisher; every other record keeps its bytes.
    pub(crate) fn publish_owner(
        self,
        native: &slide_order::Snapshot,
        owner: slide_order::records::Owner,
        before: Vec<u8>,
        after: Vec<u8>,
    ) -> Result<Published, Failure> {
        if before == after {
            return self.unchanged();
        }
        if self.is_signed()? {
            return Err(Failure::Signed);
        }
        let published = slide_order::records::publish(
            native,
            &[slide_order::records::Rewrite {
                owner,
                before: Arc::from(before),
                after: Arc::from(after),
            }],
        )?;
        self.publish(published.bytes().to_vec())
    }

    /// The exact no-op publication: the source package unchanged.
    pub(crate) fn unchanged(self) -> Result<Published, Failure> {
        let bytes = self.shared_bytes()?;
        let digest = self.digest()?;
        Ok(Published {
            patch: BytePatch::noop(bytes.len(), digest),
            package: self.package,
            diagnostics: Diagnostics::UNCHANGED,
        })
    }
}

/// How a publication seeds the published package's digest cache.
enum Seed {
    /// The target's block list, derived at commit.
    Blocks(Arc<BlockDigests>),
    /// The recipe of the target's block list, derived on first use; the
    /// root is the patch's target root.
    Derived(Recipe),
}

/// Replay an exact byte patch onto the package it was made from.
///
/// The receiving package's block root (cached, seeded, or computed once)
/// must equal the patch source's. The rebuilt package is seeded with the
/// patch's target root and the recipe of its block list, so the replay
/// hashes no block. The public family patch is opaque and is created only by
/// the closed publishers above; there is no public byte/JSON decoder that
/// could supply an arbitrary target to this path.
pub(crate) fn apply(package: &Package, patch: &BytePatch) -> Result<Published, Failure> {
    if package.is_encrypted() {
        return Err(Failure::Package(PackageError::EncryptionPolicy));
    }
    // Keep the one exact source image in its original Vec. Replay only borrows
    // it until it has rebuilt the target, so it does not need to create a
    // retained `Base` or an additional source `Arc`.
    let bytes = package.to_encrypted_bytes()?;
    let digest = package.exact().root(&bytes);
    if !patch.is_source(bytes.len(), &digest) {
        return Err(Failure::PatchConflict);
    }
    if patch.is_noop() {
        return Ok(Published {
            patch: BytePatch::noop(bytes.len(), digest),
            package: package.clone(),
            diagnostics: Diagnostics::UNCHANGED,
        });
    }
    admit_output_length(package, patch.target_len())?;
    let recipe = package.exact().child(&bytes, patch.spans());
    let target = patch.rebuild(bytes).ok_or(Failure::PatchConflict)?;
    publish_patched(package, target, patch.clone(), Seed::Derived(recipe))
}

/// Authenticate the source package before a changed target is admitted or
/// opened. The source-version fence is required for positional packages even
/// when the signature answer is already cached.
fn source_signed(package: &Package) -> Result<bool, Failure> {
    package.source_version()?;
    Ok(package
        .exact()
        .signed(|| package.contains_signature_component())?)
}

/// Admit a candidate package before a replay allocates its rebuilt bytes.
/// Preserve the established signed-source precedence for an over-limit
/// target: changed signed packages are refused before target publication.
fn admit_output_length(package: &Package, observed: usize) -> Result<(), Failure> {
    let maximum = package.source_limits().max_input_bytes();
    if observed > maximum {
        if source_signed(package)? {
            return Err(Failure::Signed);
        }
        return Err(Failure::Package(PackageError::Limit {
            kind: crate::SourceLimitKind::InputBytes,
            observed,
            maximum,
        }));
    }
    Ok(())
}

/// Publish a rebuilt target against an already-authenticated source owner.
/// This borrowed form lets replay retain the source in its one `Vec` instead
/// of wrapping it in a short-lived `Base` and `Arc<[u8]>`.
fn publish_patched(
    source: &Package,
    after: Vec<u8>,
    patch: BytePatch,
    seed: Seed,
) -> Result<Published, Failure> {
    let signed = source_signed(source)?;
    if signed {
        return Err(Failure::Signed);
    }
    let package = Package::from_owned_bytes_with_options(after, open_options(source))?;
    match seed {
        Seed::Blocks(blocks) => package.exact().seed_blocks(blocks),
        Seed::Derived(recipe) => package.exact().seed_derived(patch.target.digest, recipe),
    }
    // The closed publisher set above proves that an unsigned source stays
    // unsigned. Do not turn this into a target directory scan: that scan is
    // precisely the replay cost this seed removes.
    debug_assert!(!signed);
    package.exact().seed_unsigned();
    Ok(Published {
        package,
        patch,
        diagnostics: Diagnostics::CHANGED,
    })
}

fn open_options(package: &Package) -> OpenOptions<'static> {
    let options = OpenOptions::new()
        .source_limits(package.source_limits())
        .leniency(package.leniency());
    #[cfg(feature = "vba-inspection")]
    let options = options.vba_project_limits(package.vba_project_limits());
    options
}

#[cfg(test)]
fn contains_signature(bytes: &[u8]) -> Result<bool, Failure> {
    Ok(crate::raw::Reader::from_reader(Cursor::new(bytes))?.contains_signature_component()?)
}

/// Declare the canonical family types of one scoped family in the calling
/// module: `Diagnostics`, `LimitKind`, `Error`, `Result`, `Patch` and
/// `Commit`.
macro_rules! family {
    ($noun:literal) => {
        #[doc = concat!("Content-free diagnostics of one ", $noun, " commit.")]
        pub use litchi_core::edit::Diagnostics;

        #[doc = concat!(
                            "A published ", $noun, " edit: the new package, its exact patch and ",
                            "diagnostics."
                        )]
        pub type Commit = litchi_core::edit::Commit<$crate::Package, Patch>;

        #[doc = concat!(
                            "A ", $noun, " edit limit dimension.\n\n",
                            "The records bound their own sizes and the republished package is ",
                            "admitted under the source package's [`SourceLimits`]",
                            "(crate::SourceLimits), so no caller-selected dimension exists."
                        )]
        #[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
        #[non_exhaustive]
        pub enum LimitKind {}

        #[doc = concat!(
                            "A ", $noun, " edit failure. Messages carry no authored content."
                        )]
        #[derive(Debug)]
        #[non_exhaustive]
        pub enum Error {
            /// The source package could not be read, or the edited bytes
            /// could not be republished as a PPT package.
            Package($crate::package::Error),
            /// A well-formed request this route refuses, classified by kind.
            Refused(litchi_core::edit::Refusal),
            /// The patch was made from a different source package.
            PatchConflict,
            /// A changed write would invalidate the package signature.
            Signed,
        }

        impl ::std::fmt::Display for Error {
            fn fmt(&self, formatter: &mut ::std::fmt::Formatter<'_>) -> ::std::fmt::Result {
                match self {
                    Self::Package(_) => {
                        formatter.write_str(concat!("PPT ", $noun, " edit failed in the package"))
                    },
                    Self::Refused(refusal) => ::std::fmt::Display::fmt(refusal, formatter),
                    Self::PatchConflict => formatter.write_str(concat!(
                        "PPT ",
                        $noun,
                        " patch source does not match the package receiving it"
                    )),
                    Self::Signed => formatter.write_str(concat!(
                        "changed PPT ",
                        $noun,
                        " writes require explicit signature removal"
                    )),
                }
            }
        }

        impl ::std::error::Error for Error {
            fn source(&self) -> Option<&(dyn ::std::error::Error + 'static)> {
                match self {
                    Self::Package(error) => Some(error),
                    Self::Refused(_) | Self::PatchConflict | Self::Signed => None,
                }
            }
        }

        impl From<$crate::package::Error> for Error {
            fn from(error: $crate::package::Error) -> Self {
                Self::Package(error)
            }
        }

        impl From<$crate::scoped::Failure> for Error {
            fn from(failure: $crate::scoped::Failure) -> Self {
                match failure {
                    $crate::scoped::Failure::Package(error) => Self::Package(error),
                    $crate::scoped::Failure::Refused(refusal) => Self::Refused(refusal),
                    $crate::scoped::Failure::PatchConflict => Self::PatchConflict,
                    $crate::scoped::Failure::Signed => Self::Signed,
                }
            }
        }

        #[doc = concat!("Result type of the ", $noun, " family.")]
        pub type Result<T> = ::std::result::Result<T, Error>;

        #[doc = concat!(
                            "An exact, reversible ", $noun, " patch: the changed byte ranges ",
                            "between two packages, bound to the exact source package (length ",
                            "and block root, record 0362) it was made from."
                        )]
        #[derive(Clone, PartialEq, Eq)]
        pub struct Patch($crate::scoped::BytePatch);

        impl ::std::fmt::Debug for Patch {
            fn fmt(&self, formatter: &mut ::std::fmt::Formatter<'_>) -> ::std::fmt::Result {
                ::std::fmt::Debug::fmt(&self.0, formatter)
            }
        }

        impl Patch {
            /// Whether applying the patch changes nothing.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                self.0.is_noop()
            }

            /// The patch that undoes this one, applicable to its result.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self(self.0.inverse())
            }

            /// The byte length of the package the patch applies to.
            #[must_use]
            pub fn source_len(&self) -> usize {
                self.0.source_len()
            }

            /// The byte length of the package the patch produces.
            #[must_use]
            pub fn target_len(&self) -> usize {
                self.0.target_len()
            }

            /// The number of package bytes the patch's changed ranges span;
            /// zero for a no-op.
            #[must_use]
            pub fn changed_bytes(&self) -> usize {
                self.0.changed_bytes()
            }
        }

        #[allow(
            dead_code,
            reason = "every family publishes through these helpers; a family may use a subset"
        )]
        fn published(published: $crate::scoped::Published) -> Commit {
            Commit::new(
                published.package,
                Patch(published.patch),
                published.diagnostics,
            )
        }

        #[allow(
            dead_code,
            reason = "every family replays through this helper; a family may use a subset"
        )]
        fn apply_patch(package: &$crate::Package, patch: &Patch) -> Result<Commit> {
            Ok(published($crate::scoped::apply(package, &patch.0)?))
        }
    };
}

pub(crate) use family;

/// `package` with a digital-signature storage added (`_xmlsignatures`
/// with one stream), the shape a signed deck has, for the families'
/// signed-package tests.
#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "a test fixture builder panics on failure by design"
)]
pub(crate) fn signed(package: &Package) -> Package {
    use litchi_cfb::{OleFile, OleWriter};

    let mut source = OleFile::open(
        Cursor::new(package.to_bytes().unwrap()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let mut writer = OleWriter::new();
    for path in source.list_streams() {
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        let data = source.open_stream(&refs).unwrap();
        writer.create_stream(&refs, &data).unwrap();
    }
    writer
        .create_stream(&["_xmlsignatures", "origin.sigs"], b"signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let signed = Package::from_bytes(&output.into_inner()).unwrap();
    assert!(contains_signature(&signed.to_bytes().unwrap()).unwrap());
    signed
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "a test fixture builder panics on failure by design"
)]
fn signature_package(package: &Package, marker: &str) -> Package {
    use litchi_cfb::{OleFile, OleWriter};

    let mut source = OleFile::open(
        Cursor::new(package.to_bytes().unwrap()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let mut writer = OleWriter::new();
    for path in source.list_streams() {
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        let data = source.open_stream(&refs).unwrap();
        writer.create_stream(&refs, &data).unwrap();
    }
    writer
        .create_stream(&[marker, "origin.sigs"], b"signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    Package::from_bytes(&output.into_inner()).unwrap()
}

#[cfg(test)]
#[allow(
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "CFB fixture mutation is bounded by the writer output"
)]
fn cfb_read_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

#[cfg(test)]
#[allow(
    clippy::indexing_slicing,
    reason = "CFB fixture mutation is bounded by the writer output"
)]
fn cfb_directory_offsets(bytes: &[u8]) -> Vec<usize> {
    let sector_shift = u16::from_le_bytes(bytes[0x1e..0x20].try_into().unwrap());
    let sector_size = 1usize << sector_shift;
    let mut offsets = Vec::new();
    let mut directory_sector = cfb_read_u32(bytes, 0x30);
    while directory_sector != u32::MAX && directory_sector != u32::MAX - 1 {
        let base = (directory_sector as usize + 1) * sector_size;
        offsets.extend((0..sector_size / 128).map(|index| base + index * 128));
        directory_sector = cfb_read_u32(
            bytes,
            (cfb_read_u32(bytes, 0x4c) as usize + 1) * sector_size + directory_sector as usize * 4,
        );
    }
    offsets
}

#[cfg(test)]
#[allow(
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "CFB fixture mutation is bounded by the writer output"
)]
fn cfb_entry_name(bytes: &[u8], offset: usize) -> String {
    let name_len = usize::from(u16::from_le_bytes(
        bytes[offset + 64..offset + 66].try_into().unwrap(),
    ));
    let units = bytes[offset..offset + name_len.saturating_sub(2)]
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes(pair.try_into().unwrap()))
        .collect::<Vec<_>>();
    String::from_utf16(&units).unwrap()
}

#[cfg(test)]
#[allow(
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "CFB fixture mutation is bounded by the writer output"
)]
fn rename_cfb_root(bytes: &mut [u8], name: &str) {
    let offset = *cfb_directory_offsets(bytes).first().unwrap();
    let mut encoded = name.encode_utf16().collect::<Vec<_>>();
    encoded.push(0);
    assert!(encoded.len() <= 32);
    bytes[offset..offset + 64].fill(0);
    for (index, unit) in encoded.iter().enumerate() {
        bytes[offset + index * 2..offset + index * 2 + 2].copy_from_slice(&unit.to_le_bytes());
    }
    bytes[offset + 64..offset + 66]
        .copy_from_slice(&u16::try_from(encoded.len() * 2).unwrap().to_le_bytes());
}

#[cfg(test)]
#[allow(
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "CFB fixture mutation is bounded by the writer output"
)]
fn detach_cfb_leaf(bytes: &mut [u8], name: &str) {
    let offsets = cfb_directory_offsets(bytes);
    let (sid, offset) = offsets
        .iter()
        .enumerate()
        .find(|&(_, &offset)| bytes[offset + 66] != 0 && cfb_entry_name(bytes, offset) == name)
        .map(|(sid, &offset)| (u32::try_from(sid).unwrap(), offset))
        .unwrap();
    // The marker sorts after the generated PowerPoint entries, so it has no
    // right subtree. Reattach its lower subtree at the parent link to keep
    // every non-marker stream reachable while leaving the marker physical
    // directory slot in place as an orphan.
    assert_eq!(cfb_read_u32(bytes, offset + 72), u32::MAX);
    let replacement = cfb_read_u32(bytes, offset + 68);
    let mut detached = false;
    for parent in offsets {
        for link in [parent + 68, parent + 72, parent + 76] {
            if cfb_read_u32(bytes, link) == sid {
                bytes[link..link + 4].copy_from_slice(&replacement.to_le_bytes());
                detached = true;
                break;
            }
        }
        if detached {
            break;
        }
    }
    assert!(detached, "{name} has no reachable parent link");
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::indexing_slicing,
    reason = "byte-patch unit tests use panic-driven assertions"
)]
mod tests {
    use super::{BLOCK, Base, BlockDigests, BytePatch, Endpoint};
    use crate::package::{Error, OpenOptions, Package, SourceLimits};
    use litchi_cfb::OleError;
    use litchi_core::{ReadAt, SourceVersion};
    use std::sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    };

    fn root(bytes: &[u8]) -> super::Digest {
        BlockDigests::of(bytes).root()
    }

    #[test]
    fn base_capture_defers_the_exact_source_copy_until_requested() {
        let package = Package::new().unwrap();
        let base = Base::capture(&package).unwrap();
        assert!(base.bytes.value.get().is_none());
        drop(base);

        let base = Base::capture(&package).unwrap();
        let expected = package.to_encrypted_bytes().unwrap();
        assert_eq!(base.bytes().unwrap(), expected.as_slice());
        assert!(base.bytes.value.get().is_some());
    }

    impl BytePatch {
        fn diff(before: &[u8], after: &[u8]) -> Self {
            let ranges = Self::ranges(before, after);
            Self {
                source: Endpoint::new(before.len(), root(before)),
                target: Endpoint::new(after.len(), root(after)),
                block_spans: Arc::from(super::spans(&ranges)),
                ranges,
                inverted: false,
            }
        }

        /// What `apply` does with the package bytes: check the source,
        /// rebuild, and check the rebuilt target's digest.
        fn target_of(&self, source: &[u8]) -> Option<Vec<u8>> {
            if !self.is_source(source.len(), &root(source)) {
                return None;
            }
            let target = self.rebuild(source.to_vec())?;
            (root(&target) == self.target.digest).then_some(target)
        }
    }

    fn image(len: usize, seed: u8) -> Vec<u8> {
        (0..len)
            .map(|index| (index as u8).wrapping_mul(31).wrapping_add(seed))
            .collect()
    }

    struct VersionedSource {
        bytes: Vec<u8>,
        revision: AtomicU64,
        flip_after_versions: AtomicU64,
    }

    impl VersionedSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                revision: AtomicU64::new(0),
                flip_after_versions: AtomicU64::new(0),
            }
        }

        fn change(&self) {
            self.revision.fetch_add(1, Ordering::AcqRel);
        }

        fn change_before_version_call(&self, calls: u64) {
            self.flip_after_versions.store(calls, Ordering::Release);
        }
    }

    impl ReadAt for VersionedSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            let remaining = self.flip_after_versions.load(Ordering::Acquire);
            if remaining == 1 {
                if self
                    .flip_after_versions
                    .compare_exchange(1, 0, Ordering::AcqRel, Ordering::Acquire)
                    .is_ok()
                {
                    self.revision.fetch_add(1, Ordering::AcqRel);
                }
            } else if remaining > 1 {
                self.flip_after_versions.fetch_sub(1, Ordering::AcqRel);
            }
            let revision = self.revision.load(Ordering::Acquire);
            Ok(SourceVersion::new(0x5050, revision))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            let offset = usize::try_from(offset)
                .map_err(|_| std::io::Error::from(std::io::ErrorKind::InvalidInput))?;
            let input = self.bytes.get(offset..).unwrap_or_default();
            let count = input.len().min(output.len());
            output[..count].copy_from_slice(&input[..count]);
            Ok(count)
        }
    }

    #[test]
    fn signed_cache_rechecks_a_changed_positional_source() {
        let source_package = Package::new().unwrap();
        let source = Arc::new(VersionedSource::new(
            source_package.to_encrypted_bytes().unwrap(),
        ));
        let package = Package::from_read_at(source.clone()).unwrap();
        let base = Base::capture(&package).unwrap();

        assert!(!base.is_signed().unwrap());
        source.change();
        assert!(matches!(
            base.is_signed(),
            Err(super::Failure::Package(Error::Ole(
                OleError::SourceChanged { .. }
            )))
        ));
    }

    #[test]
    fn lazy_source_capture_rejects_a_change_before_first_materialization() {
        let source_package = Package::new().unwrap();
        let source = Arc::new(VersionedSource::new(
            source_package.to_encrypted_bytes().unwrap(),
        ));
        let package = Package::from_read_at(source.clone()).unwrap();
        let base = Base::capture(&package).unwrap();
        source.change();

        assert!(matches!(
            base.bytes(),
            Err(super::Failure::Package(Error::Ole(
                OleError::SourceChanged { .. }
            )))
        ));
    }

    #[test]
    fn lazy_source_capture_rechecks_a_change_after_materialization() {
        let source_package = Package::new().unwrap();
        let source = Arc::new(VersionedSource::new(
            source_package.to_encrypted_bytes().unwrap(),
        ));
        let package = Package::from_read_at(source.clone()).unwrap();
        let base = Base::capture(&package).unwrap();
        assert!(!base.bytes().unwrap().is_empty());
        source.change();

        assert!(matches!(
            base.bytes(),
            Err(super::Failure::Package(Error::Ole(
                OleError::SourceChanged { .. }
            )))
        ));
    }

    #[test]
    fn native_projection_rechecks_a_cached_positional_source() {
        let source_package = Package::new().unwrap();
        let source = Arc::new(VersionedSource::new(
            source_package.to_encrypted_bytes().unwrap(),
        ));
        let package = Package::from_read_at(source.clone()).unwrap();
        let base = Base::capture(&package).unwrap();
        base.bytes().unwrap();
        source.change();

        assert!(matches!(
            base.native(),
            Err(super::Failure::Package(Error::Ole(
                OleError::SourceChanged { .. }
            )))
        ));
    }

    #[test]
    fn native_projection_fences_a_source_change_during_construction() {
        let source_package = Package::new().unwrap();
        let source = Arc::new(VersionedSource::new(
            source_package.to_encrypted_bytes().unwrap(),
        ));
        let package = Package::from_read_at(source.clone()).unwrap();
        let base = Base::capture(&package).unwrap();
        base.bytes().unwrap();

        // Arm the source immediately before the final Base fence. The
        // package's own presentation fence has already accepted the same
        // source revision, so this isolates the owner-level check.
        source.change_before_version_call(9);
        assert!(matches!(
            base.native(),
            Err(super::Failure::Package(Error::Ole(
                OleError::SourceChanged { .. }
            )))
        ));

        // The same transition does not make native construction fail by
        // itself: without Base::native's final fence, the package's own
        // presentation fence observes the new revision and succeeds.
        let control_source = Arc::new(VersionedSource::new(
            source_package.to_encrypted_bytes().unwrap(),
        ));
        let control_package = Package::from_read_at(control_source.clone()).unwrap();
        let control = Base::capture(&control_package).unwrap();
        control.bytes().unwrap();
        control_source.change_before_version_call(9);
        let bytes = control.shared_bytes().unwrap();
        let control_snapshot = crate::slide_order::Snapshot::from_package_with_options(
            control.package(),
            bytes,
            control.open_options(),
        );
        assert!(control_snapshot.is_ok());
    }

    #[test]
    fn changed_owned_publication_seeds_its_unsigned_signature_result() {
        let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/unified-ops/0294/video.ppt");
        let source = Package::from_bytes(&std::fs::read(path).unwrap()).unwrap();
        let mut edit = source.edit_external_media().unwrap();
        edit.set_path(1, Some("seeded-signature-cache.avi".into()))
            .unwrap();
        let published = edit.commit().unwrap();

        let mut scans = 0;
        let signed = published
            .package()
            .exact()
            .signed(|| {
                scans += 1;
                Ok::<bool, ()>(true)
            })
            .unwrap();
        assert!(!signed);
        assert_eq!(scans, 0);
    }

    #[test]
    fn replay_works_on_an_independent_copy_and_fences_positional_sources() {
        let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/unified-ops/0294/video.ppt");
        let bytes = std::fs::read(path).unwrap();
        let source = Package::from_bytes(&bytes).unwrap();
        let mut edit = source.edit_external_media().unwrap();
        edit.set_path(1, Some("independent-copy-check.avi".into()))
            .unwrap();
        let commit = edit.commit().unwrap();

        let independent = Package::from_bytes(&bytes).unwrap();
        let replayed = independent.apply_external_media(commit.patch()).unwrap();
        assert_eq!(
            replayed.package().to_bytes().unwrap(),
            commit.package().to_bytes().unwrap()
        );
        let restored = replayed
            .package()
            .apply_external_media(&commit.patch().inverse())
            .unwrap();
        assert_eq!(restored.package().to_bytes().unwrap(), bytes);

        let versioned = Arc::new(VersionedSource::new(bytes));
        let positional = Package::from_read_at(versioned.clone()).unwrap();
        versioned.change();
        assert!(matches!(
            positional.apply_external_media(commit.patch()),
            Err(crate::external_media::Error::Package(Error::Ole(
                OleError::SourceChanged { .. }
            )))
        ));
    }

    #[test]
    fn replay_admits_a_growing_target_under_the_recipient_limit() {
        let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/unified-ops/0294/video.ppt");
        let source_bytes = std::fs::read(path).unwrap();
        let source = Package::from_bytes(&source_bytes).unwrap();
        let mut edit = source.edit_external_media().unwrap();
        edit.set_path(1, Some(format!("{}.avi", "long-path".repeat(1024))))
            .unwrap();
        let commit = edit.commit().unwrap();
        assert!(commit.patch().target_len() > source_bytes.len());

        let limits = SourceLimits::builder()
            .max_input_bytes(source_bytes.len())
            .unwrap()
            .build()
            .unwrap();
        let recipient = Package::from_bytes_with_options(
            &source_bytes,
            OpenOptions::new().source_limits(limits),
        )
        .unwrap();
        assert!(matches!(
            recipient.apply_external_media(commit.patch()),
            Err(crate::external_media::Error::Package(Error::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                observed,
                maximum,
            })) if observed == commit.patch().target_len() && maximum == source_bytes.len()
        ));
    }

    #[test]
    fn signature_scan_follows_reachable_nested_case_insensitive_entries() {
        let package = Package::new().unwrap();
        assert!(!package.contains_signature_component().unwrap());

        let mixed_case = super::signature_package(&package, "_XMLSIGNATURES");
        assert!(mixed_case.contains_signature_component().unwrap());

        let mut renamed_root_bytes = package.to_bytes().unwrap();
        super::rename_cfb_root(&mut renamed_root_bytes, "_signatures");
        let raw = crate::raw::Reader::from_reader(std::io::Cursor::new(renamed_root_bytes.clone()))
            .unwrap();
        assert!(!raw.contains_signature_component().unwrap());
        let renamed_root = Package::from_bytes(&renamed_root_bytes).unwrap();
        assert!(!renamed_root.contains_signature_component().unwrap());

        let mut orphan_bytes = mixed_case.to_bytes().unwrap();
        super::detach_cfb_leaf(&mut orphan_bytes, "_XMLSIGNATURES");
        let raw =
            crate::raw::Reader::from_reader(std::io::Cursor::new(orphan_bytes.clone())).unwrap();
        assert!(!raw.contains_signature_component().unwrap());
        let orphan = Package::from_bytes(&orphan_bytes).unwrap();
        assert!(!orphan.contains_signature_component().unwrap());

        let nested = super::signed(&package);
        assert!(nested.contains_signature_component().unwrap());

        let strict_events = SourceLimits::builder()
            .max_events(1)
            .unwrap()
            .build()
            .unwrap();
        let bounded = crate::raw::Reader::from_reader_with_options(
            std::io::Cursor::new(nested.to_bytes().unwrap()),
            OpenOptions::new().source_limits(strict_events),
        )
        .unwrap();
        assert!(matches!(
            bounded.contains_signature_component(),
            Err(Error::Ole(OleError::LimitExceeded {
                resource: "directory entries",
                maximum: 1,
                ..
            }))
        ));
        // The canonical owner admits its presentation at open, so this
        // tighter profile fails before a signature projection is available.
        assert!(matches!(
            Package::from_bytes_with_options(
                &nested.to_bytes().unwrap(),
                OpenOptions::new().source_limits(strict_events),
            ),
            Err(Error::Limit {
                kind: crate::SourceLimitKind::Events,
                observed: 2,
                maximum: 1,
            })
        ));
    }

    #[test]
    fn compact_span_publication_admits_length_before_constructing_ranges() {
        let source = Package::new().unwrap();
        let bytes = source.to_bytes().unwrap();
        let maximum = bytes.len();
        let limits = SourceLimits::builder()
            .max_input_bytes(maximum)
            .unwrap()
            .build()
            .unwrap();
        let package =
            Package::from_bytes_with_options(&bytes, OpenOptions::new().source_limits(limits))
                .unwrap();
        let base = Base::capture(&package).unwrap();
        base.admit_output_length(maximum).unwrap();
        for observed in [maximum + 1, usize::MAX] {
            assert!(matches!(
                base.admit_output_length(observed),
                Err(super::Failure::Package(Error::Limit {
                    kind: crate::SourceLimitKind::InputBytes,
                    observed: actual,
                    maximum: ceiling,
                })) if actual == observed && ceiling == maximum
            ));
        }
        let mut oversized = bytes;
        oversized.push(0);
        assert!(matches!(
            base.publish(oversized),
            Err(super::Failure::Package(Error::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                observed,
                maximum: ceiling,
            })) if observed == maximum + 1 && ceiling == maximum
        ));

        let signed = super::signed(&source);
        let signed_base = Base::capture(&signed).unwrap();
        assert!(matches!(
            signed_base.admit_output_length(usize::MAX),
            Err(super::Failure::Signed)
        ));
        let versioned = Arc::new(VersionedSource::new(source.to_bytes().unwrap()));
        let package = Package::from_read_at(versioned.clone()).unwrap();
        let base = Base::capture(&package).unwrap();
        versioned.change();
        assert!(matches!(
            base.admit_output_length(usize::MAX),
            Err(super::Failure::Package(Error::Ole(
                OleError::SourceChanged { .. }
            )))
        ));
    }

    #[test]
    fn signature_projection_bounds_a_large_admitted_directory() {
        use litchi_cfb::{OleFile, OleWriter};

        let package = Package::new().unwrap();
        let mut source = OleFile::open(
            std::io::Cursor::new(package.to_bytes().unwrap()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        let mut writer = OleWriter::new();
        for path in source.list_streams() {
            let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
            writer
                .create_stream(&refs, &source.open_stream(&refs).unwrap())
                .unwrap();
        }
        for index in 0..1025 {
            writer
                .create_stream(&[&format!("Unused{index:04}")], &[])
                .unwrap();
        }
        let mut output = std::io::Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        let limits = SourceLimits::builder()
            .max_events(1024)
            .unwrap()
            .build()
            .unwrap();
        let package = Package::from_bytes_with_options(
            &output.into_inner(),
            OpenOptions::new().source_limits(limits),
        )
        .unwrap();
        assert!(matches!(
            package.contains_signature_component(),
            Err(Error::Ole(OleError::LimitExceeded {
                resource: "directory entries",
                observed: 1025,
                maximum: 1024,
            }))
        ));
    }

    #[test]
    fn unsigned_publication_does_not_reach_an_orphan_signature_component() {
        let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/unified-ops/0294/video.ppt");
        let source = Package::from_bytes(&std::fs::read(path).unwrap()).unwrap();
        let signed = super::signature_package(&source, "_XMLSIGNATURES");
        let mut bytes = signed.to_bytes().unwrap();
        super::detach_cfb_leaf(&mut bytes, "_XMLSIGNATURES");
        let orphan = Package::from_bytes(&bytes).unwrap();
        assert!(!orphan.contains_signature_component().unwrap());

        let mut edit = orphan.edit_external_media().unwrap();
        edit.set_path(1, Some("orphan-marker-check.avi".into()))
            .unwrap();
        let published = edit.commit().unwrap();
        assert!(!published.package().contains_signature_component().unwrap());

        let mut scans = 0;
        let cached = published
            .package()
            .exact()
            .signed(|| {
                scans += 1;
                Ok::<bool, ()>(true)
            })
            .unwrap();
        assert!(!cached);
        assert_eq!(scans, 0);
    }

    /// `patch` maps `before` to `after` exactly, its inverse maps back, and
    /// neither applies to the other's source.
    fn assert_exact(before: &[u8], after: &[u8]) -> BytePatch {
        let patch = BytePatch::diff(before, after);
        assert_eq!(patch.source_len(), before.len());
        assert_eq!(patch.target_len(), after.len());
        assert_eq!(patch.target_of(before).unwrap(), after);
        let inverse = patch.inverse();
        assert_eq!(inverse.target_of(after).unwrap(), before);
        assert_eq!(inverse.changed_bytes(), patch.changed_bytes());
        assert_eq!(inverse.inverse(), patch);
        assert_eq!(inverse, BytePatch::diff(after, before));
        if before != after {
            assert!(patch.target_of(after).is_none());
            assert!(inverse.target_of(before).is_none());
        }
        patch
    }

    #[test]
    fn a_no_op_diff_has_no_ranges_and_applies_only_to_its_source() {
        let bytes = image(5 * BLOCK + 17, 3);
        let patch = assert_exact(&bytes, &bytes);
        assert!(patch.is_noop());
        assert_eq!(patch.changed_bytes(), 0);
        assert_eq!(patch, BytePatch::noop(bytes.len(), root(&bytes)));
        let mut other = bytes.clone();
        other[BLOCK] ^= 1;
        assert!(patch.target_of(&other).is_none());
        assert!(patch.target_of(&bytes[..bytes.len() - 1]).is_none());
    }

    #[test]
    fn a_mid_change_spans_only_its_coalesced_blocks() {
        let before = image(8 * BLOCK, 7);
        let mut after = before.clone();
        after[2 * BLOCK + 5] ^= 0xff;
        after[3 * BLOCK + 9] ^= 0xff;
        after[6 * BLOCK] ^= 0xff;
        let patch = assert_exact(&before, &after);
        assert_eq!(patch.ranges.len(), 2);
        assert_eq!(patch.ranges[0].offset, 2 * BLOCK);
        assert_eq!(patch.ranges[0].after.len(), 2 * BLOCK);
        assert_eq!(patch.ranges[1].offset, 6 * BLOCK);
        assert_eq!(patch.changed_bytes(), 3 * BLOCK);
    }

    #[test]
    fn growth_and_shrinkage_belong_to_the_last_range() {
        let before = image(4 * BLOCK + 100, 11);
        let mut grown = before.clone();
        grown.extend_from_slice(&image(3 * BLOCK + 1, 13));
        let patch = assert_exact(&before, &grown);
        assert_eq!(patch.ranges.len(), 1);
        assert_eq!(patch.ranges[0].offset, 4 * BLOCK);
        assert_eq!(patch.ranges[0].before.len(), 100);
        assert_eq!(patch.ranges[0].after.len(), grown.len() - 4 * BLOCK);

        let mut changed_and_grown = grown.clone();
        changed_and_grown[0] ^= 1;
        let patch = assert_exact(&before, &changed_and_grown);
        assert_eq!(patch.ranges.len(), 2);

        let shrunk = before[..BLOCK + 3].to_vec();
        let patch = assert_exact(&before, &shrunk);
        assert_eq!(patch.ranges.len(), 1);
        assert_eq!(patch.ranges[0].offset, BLOCK);
        assert_eq!(patch.ranges[0].after.len(), 3);
        assert_eq!(patch.ranges[0].before.len(), before.len() - BLOCK);

        // A shrink to an exact block boundary leaves an empty target side.
        let patch = assert_exact(&before, &before[..2 * BLOCK]);
        assert_eq!(patch.ranges[0].after.len(), 0);
        assert_exact(&before, &[]);
        assert_exact(&[], &before);
    }

    #[test]
    fn rebuild_reuses_an_owned_source_buffer_when_lengths_fit() {
        let before = image(3 * BLOCK + 11, 19);
        let mut after = before.clone();
        after[BLOCK + 3] ^= 0x5a;
        let patch = BytePatch::diff(&before, &after);
        let source = before.clone();
        let source_pointer = source.as_ptr();
        let rebuilt = patch.rebuild(source).unwrap();
        assert_eq!(rebuilt.as_ptr(), source_pointer);
        assert_eq!(rebuilt, after);
    }

    #[test]
    fn a_patch_refuses_a_same_length_package_that_differs_outside_its_ranges() {
        let before = image(6 * BLOCK, 17);
        let mut after = before.clone();
        after[BLOCK] ^= 1;
        let patch = assert_exact(&before, &after);
        let mut wrong = before.clone();
        wrong[5 * BLOCK] ^= 1;
        assert!(patch.target_of(&wrong).is_none());
    }
}
