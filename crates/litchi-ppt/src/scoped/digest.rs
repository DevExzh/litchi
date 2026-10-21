//! Incremental exact-package digests of the scoped families (record 0357b).
//!
//! A scoped patch binds its source and target package by length and a
//! *block root*: the SHA-256 over the image length and a 128-bit digest of
//! every 512-byte block (the diff's block size, one OLE sector). The block
//! digest is SipHash-2-4 with a 128-bit output under a key drawn once per
//! process from the standard library's randomly keyed hasher state. Under
//! that keyed model, a caller without the key cannot efficiently target a
//! collision; the key reduces adversarial targeting rather than making
//! collisions impossible. Two different blocks still collide by chance with
//! probability 2^-128 (2^-64 over all pairs of a 2^32-block history). Two
//! images with the same root therefore have the same length and, short of such
//! a collision, the same bytes. Scoped patches are process-local, so a
//! per-process key never has to match across processes. SipHash replaced the
//! earlier per-block SHA-256; the keyed BLAKE2b replacement was measured and
//! rejected for its first-commit regression (record 0362).
//!
//! The block list lets a commit derive the target's digests from the
//! source's: only the blocks a patch changes or appends are hashed again,
//! and the root is one hash over the list (16 bytes per block). Every
//! [`Package`](crate::Package) owns an [`ExactCache`]: its block list is
//! computed once, on first use, unless the publisher that produced the
//! package seeded it, so a chain of commits and applies hashes only what
//! each step changed.

use std::sync::{Arc, OnceLock};

use sha2::{Digest as _, Sha256};
use siphasher::sip128::SipHasher24;

use super::BLOCK;

/// A SHA-256 digest (a block root).
pub(crate) type Digest = [u8; 32];

/// A keyed 128-bit block digest.
type BlockDigest = [u8; 16];

/// A compact block-index span retained by a patch and its lazy recipe.
///
/// PPT source limits keep the admitted image below the `u32` block-index
/// ceiling. Keeping the indices compact halves the per-patch span payload
/// while preserving the exact changed-block set used by replay.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct BlockSpan {
    start: u32,
    end: u32,
}

impl BlockSpan {
    pub(crate) fn new(start: usize, end: usize) -> Self {
        Self {
            start: u32::try_from(start).expect("PPT block span start exceeds u32"),
            end: u32::try_from(end).expect("PPT block span end exceeds u32"),
        }
    }

    fn start(self) -> usize {
        self.start as usize
    }

    fn end(self) -> usize {
        self.end as usize
    }
}

/// The domain tag hashed before a block root's length and block digests.
const ROOT_TAG: &[u8] = b"litchi-ppt/scoped/block-root/v2";

/// The process's secret block-digest hasher, initialized once from the
/// standard library's randomly keyed hasher state (seeded from the operating
/// system). Each block clones this initialized state instead of rebuilding
/// the keyed SipHash state.
fn hasher() -> SipHasher24 {
    static HASHER: OnceLock<SipHasher24> = OnceLock::new();
    *HASHER.get_or_init(|| {
        use std::hash::BuildHasher as _;
        let state = std::collections::hash_map::RandomState::new();
        SipHasher24::new_with_keys(
            state.hash_one(0x6c69_7463_6869_5f6b_u64),
            state.hash_one(0x6579_5f62_6c6f_636b_u64),
        )
    })
}

fn block_digest(hasher: &SipHasher24, bytes: &[u8]) -> BlockDigest {
    hasher.hash(bytes).as_bytes()
}

/// Block `index` of `bytes`: 512 bytes, or the shorter tail.
fn block(bytes: &[u8], index: usize) -> &[u8] {
    let start = index.saturating_mul(BLOCK).min(bytes.len());
    let stop = start.saturating_add(BLOCK).min(bytes.len());
    bytes.get(start..stop).unwrap_or_default()
}

/// The keyed digest of every 512-byte block of one exact image, with its
/// length and its root, computed once.
pub(crate) struct BlockDigests {
    len: usize,
    blocks: Box<[BlockDigest]>,
    root: OnceLock<Digest>,
}

impl BlockDigests {
    /// Hash every block of `bytes`.
    pub(crate) fn of(bytes: &[u8]) -> Self {
        let hasher = hasher();
        let blocks = bytes
            .chunks(BLOCK)
            .map(|chunk| block_digest(&hasher, chunk))
            .collect();
        Self {
            len: bytes.len(),
            blocks,
            root: OnceLock::new(),
        }
    }

    /// The block list of `bytes`, an image that equals the image this list
    /// describes in every block outside `changed` (block-index spans,
    /// sorted and disjoint). Blocks in `changed`, blocks past this image,
    /// and a block whose extent differs (a tail that grew or shrank) are
    /// hashed again; every other digest is reused.
    pub(crate) fn derive(&self, bytes: &[u8], changed: &[BlockSpan]) -> Self {
        let count = bytes.len().div_ceil(BLOCK);
        let hasher = hasher();
        let mut blocks = Vec::with_capacity(count);
        let mut spans = changed.iter().peekable();
        for index in 0..count {
            while spans.next_if(|span| span.end() <= index).is_some() {}
            let is_changed = spans.peek().is_some_and(|span| span.start() <= index);
            let data = block(bytes, index);
            let reused = (!is_changed && self.extent(index) == data.len())
                .then(|| self.blocks.get(index))
                .flatten();
            blocks.push(
                reused
                    .copied()
                    .unwrap_or_else(|| block_digest(&hasher, data)),
            );
        }
        let derived = Self {
            len: bytes.len(),
            blocks: blocks.into_boxed_slice(),
            root: OnceLock::new(),
        };
        debug_assert!(derived.blocks == Self::of(bytes).blocks);
        derived
    }

    /// The byte length of block `index` of the described image.
    fn extent(&self, index: usize) -> usize {
        self.len
            .saturating_sub(index.saturating_mul(BLOCK))
            .min(BLOCK)
    }

    /// The root: SHA-256 over a domain tag, the length (u64, little-endian)
    /// and every block digest in order.
    pub(crate) fn root(&self) -> Digest {
        *self.root.get_or_init(|| {
            let mut hasher = Sha256::new();
            hasher.update(ROOT_TAG);
            hasher.update((self.len as u64).to_le_bytes());
            for digest in &self.blocks {
                hasher.update(digest);
            }
            hasher.finalize().into()
        })
    }
}

/// How to derive a package's block list without hashing all of it: the list
/// of an ancestor image and the blocks in which the package may differ.
#[derive(Clone)]
pub(crate) struct Recipe {
    from: Arc<BlockDigests>,
    changed: Arc<[BlockSpan]>,
}

/// The union of two sorted, disjoint block-span lists, merged.
fn union(left: &[BlockSpan], right: &[BlockSpan]) -> Arc<[BlockSpan]> {
    let mut all: Vec<BlockSpan> = left.iter().chain(right).copied().collect();
    all.sort_unstable_by_key(|span| span.start);
    let mut merged: Vec<BlockSpan> = Vec::with_capacity(all.len());
    for span in all {
        match merged.last_mut() {
            Some(last) if span.start <= last.end => last.end = last.end.max(span.end),
            _ => merged.push(span),
        }
    }
    Arc::from(merged)
}

/// The digest facts of one package's exact bytes, owned by the package:
/// its block list, its root, a recipe a publisher seeded instead of the
/// list, and whether it carries a digital-signature component. Each is
/// computed at most once from the package's bytes, which the caller passes
/// (they are the package's [`to_encrypted_bytes`](crate::Package::to_encrypted_bytes)).
#[derive(Default)]
pub(crate) struct ExactCache {
    blocks: OnceLock<Arc<BlockDigests>>,
    root: OnceLock<Digest>,
    recipe: OnceLock<Recipe>,
    signed: OnceLock<bool>,
}

impl ExactCache {
    /// The block list of `bytes`: cached, derived from a seeded recipe
    /// (hashing only the recipe's changed blocks), or hashed in full.
    pub(crate) fn blocks(&self, bytes: &[u8]) -> Arc<BlockDigests> {
        Arc::clone(self.blocks.get_or_init(|| {
            Arc::new(match self.recipe.get() {
                Some(recipe) => recipe.from.derive(bytes, &recipe.changed),
                None => BlockDigests::of(bytes),
            })
        }))
    }

    /// The block root of `bytes`, cached or seeded.
    pub(crate) fn root(&self, bytes: &[u8]) -> Digest {
        *self.root.get_or_init(|| self.blocks(bytes).root())
    }

    /// Seed the block list a publisher derived for these bytes.
    pub(crate) fn seed_blocks(&self, blocks: Arc<BlockDigests>) {
        let _root = self.root.set(blocks.root());
        let _blocks = self.blocks.set(blocks);
    }

    /// Seed the root a patch proved for these bytes and the recipe of their
    /// block list, derived lazily on first use.
    pub(crate) fn seed_derived(&self, root: Digest, recipe: Recipe) {
        let _root = self.root.set(root);
        let _recipe = self.recipe.set(recipe);
    }

    /// Seed an unsigned result after a publisher has authenticated the source
    /// and built the new owned package.
    ///
    /// The caller must be one of the closed scoped publishers documented on
    /// [`super::Base::publish`]. Those publishers only rewrite admitted
    /// records or generate typed embedded-object streams; they cannot create
    /// a new reachable signature directory entry. Keeping this proof at the cache
    /// boundary avoids rescanning the admitted directory on the next replay.
    pub(crate) fn seed_unsigned(&self) {
        let _signed = self.signed.set(false);
    }

    /// The recipe of an image that equals `bytes` (this package's bytes)
    /// outside the block spans `changed`. It reuses this package's list
    /// when known, or composes with this package's own recipe (the union of
    /// both span lists), so no block is hashed here.
    pub(crate) fn child(&self, bytes: &[u8], changed: &[BlockSpan]) -> Recipe {
        if self.blocks.get().is_none()
            && let Some(recipe) = self.recipe.get()
        {
            if recipe.changed.as_ref() == changed {
                return recipe.clone();
            }
            return Recipe {
                from: Arc::clone(&recipe.from),
                changed: union(&recipe.changed, changed),
            };
        }
        Recipe {
            from: self.blocks(bytes),
            changed: Arc::from(changed),
        }
    }

    /// Whether these bytes carry a digital-signature component, computed
    /// once with `compute`.
    pub(crate) fn signed<E>(&self, compute: impl FnOnce() -> Result<bool, E>) -> Result<bool, E> {
        if let Some(signed) = self.signed.get() {
            return Ok(*signed);
        }
        let signed = compute()?;
        let _set = self.signed.set(signed);
        Ok(signed)
    }
}

#[cfg(test)]
#[allow(
    clippy::indexing_slicing,
    clippy::single_range_in_vec_init,
    reason = "digest unit tests use panic-driven assertions and block-span literals"
)]
mod tests {
    use super::{BLOCK, BlockDigests, BlockSpan, ExactCache, Recipe, union};
    use std::sync::Arc;

    fn image(len: usize, seed: u8) -> Vec<u8> {
        (0..len)
            .map(|index| (index as u8).wrapping_mul(29).wrapping_add(seed))
            .collect()
    }

    #[test]
    fn a_root_binds_every_byte_and_the_length() {
        let bytes = image(3 * BLOCK + 7, 1);
        let root = BlockDigests::of(&bytes).root();
        for index in [0, BLOCK - 1, BLOCK, 3 * BLOCK + 6] {
            let mut other = bytes.clone();
            other[index] ^= 1;
            assert_ne!(BlockDigests::of(&other).root(), root);
        }
        assert_ne!(BlockDigests::of(&bytes[..bytes.len() - 1]).root(), root);
        let mut padded = bytes.clone();
        padded.push(0);
        assert_ne!(BlockDigests::of(&padded).root(), root);
        assert_ne!(BlockDigests::of(&[]).root(), root);
    }

    #[test]
    fn compact_spans_cover_the_largest_admitted_image_without_truncation() {
        let last = crate::package::SourceLimits::MAX_INPUT_BYTES.div_ceil(BLOCK);
        let span = BlockSpan::new(last - 1, last);
        assert_eq!(span.start(), last - 1);
        assert_eq!(span.end(), last);
        assert!(u32::try_from(last).is_ok());
    }

    #[test]
    fn a_derived_list_equals_the_full_list() {
        let before = image(6 * BLOCK + 100, 3);
        let source = BlockDigests::of(&before);
        let mut after = before.clone();
        after[2 * BLOCK + 1] ^= 1;
        let derived = source.derive(&after, &[BlockSpan::new(2, 3)]);
        assert_eq!(derived.root(), BlockDigests::of(&after).root());
        // Growth past a partial tail, shrinkage, and the empty image.
        let mut grown = before.clone();
        grown.extend_from_slice(&image(2 * BLOCK, 9));
        assert_eq!(
            source.derive(&grown, &[BlockSpan::new(6, 9)]).root(),
            BlockDigests::of(&grown).root()
        );
        let shrunk = &before[..2 * BLOCK + 3];
        assert_eq!(
            source.derive(shrunk, &[BlockSpan::new(2, 7)]).root(),
            BlockDigests::of(shrunk).root()
        );
        assert_eq!(source.derive(&[], &[]).root(), BlockDigests::of(&[]).root());
    }

    #[test]
    fn a_child_of_an_unresolved_recipe_composes_the_spans() {
        let first = image(5 * BLOCK, 5);
        let mut second = first.clone();
        second[BLOCK] ^= 1;
        let mut third = second.clone();
        third[4 * BLOCK] ^= 1;
        let cache = ExactCache::default();
        let root = BlockDigests::of(&second).root();
        cache.seed_derived(
            root,
            Recipe {
                from: Arc::new(BlockDigests::of(&first)),
                changed: Arc::from([BlockSpan::new(1, 2)]),
            },
        );
        assert_eq!(cache.root(&second), root);
        let child = cache.child(&second, &[BlockSpan::new(4, 5)]);
        assert_eq!(
            &*child.changed,
            &[BlockSpan::new(1, 2), BlockSpan::new(4, 5)]
        );
        let derived = ExactCache::default();
        derived.seed_derived(BlockDigests::of(&third).root(), child);
        assert_eq!(
            derived.blocks(&third).root(),
            BlockDigests::of(&third).root()
        );
        assert_eq!(
            &*union(
                &[BlockSpan::new(0, 2), BlockSpan::new(5, 6)],
                &[BlockSpan::new(1, 3), BlockSpan::new(6, 7)]
            ),
            &[BlockSpan::new(0, 3), BlockSpan::new(5, 7)]
        );
    }

    #[test]
    fn a_child_with_the_same_spans_reuses_the_unresolved_recipe() {
        let before = image(4 * BLOCK, 17);
        let mut after = before.clone();
        after[BLOCK] ^= 1;
        let from = Arc::new(BlockDigests::of(&before));
        let changed: Arc<[BlockSpan]> = Arc::from([BlockSpan::new(1, 2)]);
        let cache = ExactCache::default();
        cache.seed_derived(
            BlockDigests::of(&after).root(),
            Recipe {
                from: Arc::clone(&from),
                changed: Arc::clone(&changed),
            },
        );

        let child = cache.child(&after, &[BlockSpan::new(1, 2)]);
        assert!(Arc::ptr_eq(&child.from, &from));
        assert!(Arc::ptr_eq(&child.changed, &changed));
    }
}
