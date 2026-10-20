//! Shared plumbing of the scoped PPT edit families (record 0357).
//!
//! Every scoped family `x` follows the record 0351 §2.3 shape:
//! `Package::edit_x(&self, ..) -> x::Edit`, `x::Edit::commit(self) ->
//! x::Commit` publishing a new [`Package`], and `Package::apply_x(&self,
//! &x::Patch) -> x::Commit`. Dropping an `Edit` discards it. The families
//! keep their record-level codecs and staging owners crate-private; this
//! module captures the source package, publishes a family's candidate bytes
//! as a new package, and replays byte patches.
//!
//! A family's patch is exact and record-sized (record 0357a): it carries the
//! length and SHA-256 of the source and target package and the coalesced
//! 512-byte blocks that differ, so `apply_x` refuses a package that is not
//! the patch's exact source ([`Failure::PatchConflict`]), rebuilds the exact
//! target, and `inverse` swaps the endpoints. An exact no-op commit
//! publishes the source package unchanged (signatures included); a changed
//! commit of a signed package is refused ([`Failure::Signed`]).

use std::io::Cursor;
use std::sync::Arc;

use litchi_core::edit::{Diagnostics, Refusal, RefusalKind};
use litchi_core::patch::BlobId;

use crate::package::{Error as PackageError, OpenOptions, Package};
use crate::slide_order;

/// One side of an exact byte patch: the package length and its SHA-256.
#[derive(Clone, PartialEq, Eq)]
struct Endpoint {
    len: usize,
    digest: BlobId,
}

impl Endpoint {
    fn new(len: usize, digest: BlobId) -> Self {
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
/// It carries both endpoints' length and SHA-256 and the coalesced 512-byte
/// blocks that differ, so it applies only to the exact source (anything else
/// is [`Failure::PatchConflict`]), reproduces the exact target, and inverts
/// by swapping the endpoints and reading each range's other side. The ranges
/// are shared between a patch, its clones and its inverses.
#[derive(Clone)]
pub(crate) struct BytePatch {
    source: Endpoint,
    target: Endpoint,
    ranges: Arc<[Range]>,
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
    /// The exact no-op patch of a package of `len` bytes with `digest`.
    fn noop(len: usize, digest: BlobId) -> Self {
        let endpoint = Endpoint::new(len, digest);
        Self {
            source: endpoint.clone(),
            target: endpoint,
            ranges: Arc::from([]),
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

    /// Whether a package of `len` bytes with `digest` is the patch's exact
    /// source.
    fn is_source(&self, len: usize, digest: &BlobId) -> bool {
        self.source.len == len && &self.source.digest == digest
    }

    /// The exact target rebuilt from the exact source bytes: copy them,
    /// resize to the target length and write every range's target side.
    /// The caller has checked [`Self::is_source`]; with the source bound by
    /// length and SHA-256 the result is the exact target by construction
    /// (debug builds re-hash it).
    fn rebuild(&self, source: &[u8]) -> Option<Vec<u8>> {
        let mut target = source.to_vec();
        target.resize(self.target.len, 0);
        for range in self.ranges.iter() {
            let (_, after) = self.sides(range);
            let stop = range.offset.checked_add(after.len())?;
            target.get_mut(range.offset..stop)?.copy_from_slice(after);
        }
        debug_assert!(BlobId::of(&target) == self.target.digest);
        Some(target)
    }
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

/// The immutable source package a scoped edit started from, with its exact
/// bytes.
#[derive(Clone)]
pub(crate) struct Base {
    package: Package,
    bytes: Arc<[u8]>,
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
        let bytes: Arc<[u8]> = Arc::from(package.to_encrypted_bytes()?);
        Ok(Self {
            package: package.clone(),
            bytes,
        })
    }

    /// The exact source package bytes.
    pub(crate) fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// The SHA-256 of the exact source bytes, cached on the package.
    fn digest(&self) -> BlobId {
        self.package.exact_digest(&self.bytes)
    }

    /// Whether the source carries a digital-signature component.
    pub(crate) fn is_signed(&self) -> Result<bool, Failure> {
        contains_signature(&self.bytes)
    }

    /// A shared handle on the exact source package bytes.
    pub(crate) fn shared_bytes(&self) -> Arc<[u8]> {
        Arc::clone(&self.bytes)
    }

    /// The options every republished package is opened with: the source
    /// package's limits and leniency.
    pub(crate) fn open_options(&self) -> OpenOptions<'static> {
        open_options(&self.package)
    }

    /// The source as the native slide-order read view, for families whose
    /// records live in the live document or a slide.
    pub(crate) fn native(&self) -> Result<slide_order::Snapshot, Failure> {
        Ok(slide_order::Snapshot::from_bytes_with_options(
            self.bytes.to_vec(),
            self.open_options(),
        )?)
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
    /// to both endpoints' length and SHA-256; the target's digest is cached
    /// on the published package, so a later edit or patch of it does not
    /// hash it again.
    pub(crate) fn publish(self, after: Vec<u8>) -> Result<Published, Failure> {
        let ranges = BytePatch::ranges(&self.bytes, &after);
        if ranges.is_empty() {
            return Ok(self.unchanged());
        }
        let target = Endpoint::new(after.len(), BlobId::of(&after));
        let patch = BytePatch {
            source: Endpoint::new(self.bytes.len(), self.digest()),
            target,
            ranges,
            inverted: false,
        };
        self.publish_patched(&after, patch)
    }

    /// Publish the exact target `after` of `patch`, made from this source.
    fn publish_patched(self, after: &[u8], patch: BytePatch) -> Result<Published, Failure> {
        if contains_signature(&self.bytes)? {
            return Err(Failure::Signed);
        }
        let package = Package::from_bytes_with_options(after, self.open_options())?;
        package.seed_exact_digest(patch.target.digest.clone());
        Ok(Published {
            package,
            patch,
            diagnostics: Diagnostics::CHANGED,
        })
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
            return Ok(self.unchanged());
        }
        if contains_signature(&self.bytes)? {
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
    pub(crate) fn unchanged(self) -> Published {
        Published {
            patch: BytePatch::noop(self.bytes.len(), self.digest()),
            package: self.package,
            diagnostics: Diagnostics::UNCHANGED,
        }
    }
}

/// Replay an exact byte patch onto the package it was made from.
pub(crate) fn apply(package: &Package, patch: &BytePatch) -> Result<Published, Failure> {
    let base = Base::capture(package)?;
    if !patch.is_source(base.bytes().len(), &base.digest()) {
        return Err(Failure::PatchConflict);
    }
    if patch.is_noop() {
        return Ok(base.unchanged());
    }
    let target = patch.rebuild(base.bytes()).ok_or(Failure::PatchConflict)?;
    base.publish_patched(&target, patch.clone())
}

fn open_options(package: &Package) -> OpenOptions<'static> {
    OpenOptions::new()
        .source_limits(package.source_limits())
        .leniency(package.leniency())
}

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
                            "and SHA-256) it was made from."
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

    let mut source = OleFile::open(Cursor::new(package.to_bytes().unwrap())).unwrap();
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
    clippy::indexing_slicing,
    reason = "byte-patch unit tests use panic-driven assertions"
)]
mod tests {
    use super::{BLOCK, BlobId, BytePatch, Endpoint};

    impl BytePatch {
        fn diff(before: &[u8], after: &[u8]) -> Self {
            Self {
                source: Endpoint::new(before.len(), BlobId::of(before)),
                target: Endpoint::new(after.len(), BlobId::of(after)),
                ranges: Self::ranges(before, after),
                inverted: false,
            }
        }

        /// What `apply` does with the package bytes: check the source,
        /// rebuild, and check the rebuilt target's digest.
        fn target_of(&self, source: &[u8]) -> Option<Vec<u8>> {
            if !self.is_source(source.len(), &BlobId::of(source)) {
                return None;
            }
            let target = self.rebuild(source)?;
            (BlobId::of(&target) == self.target.digest).then_some(target)
        }
    }

    fn image(len: usize, seed: u8) -> Vec<u8> {
        (0..len)
            .map(|index| (index as u8).wrapping_mul(31).wrapping_add(seed))
            .collect()
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
        assert_eq!(patch, BytePatch::noop(bytes.len(), BlobId::of(&bytes)));
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
