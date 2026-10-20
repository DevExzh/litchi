//! Incremental exact-package digests of the scoped families (record 0357b).
//!
//! A scoped patch binds its source and target package by length and a
//! *block root*: the SHA-256 over the image length and the SHA-256 of every
//! 512-byte block (the diff's block size, one OLE sector). Two images with
//! the same root have the same length and the same digest in every block,
//! so, short of a SHA-256 collision, they are the same bytes; the root is as
//! exact a binding as one SHA-256 over the whole image.
//!
//! The block list lets a commit derive the target's digests from the
//! source's: only the blocks a patch changes or appends are hashed again,
//! and the root is one hash over the list (32 bytes per block). Every
//! [`Package`](crate::Package) owns an [`ExactCache`]: its block list is
//! computed once, on first use, unless the publisher that produced the
//! package seeded it, so a chain of commits and applies hashes only what
//! each step changed.

use std::ops::Range;
use std::sync::{Arc, OnceLock};

use sha2::{Digest as _, Sha256};

use super::BLOCK;

/// A SHA-256 digest.
pub(crate) type Digest = [u8; 32];

/// The domain tag hashed before a block root's length and block digests.
const ROOT_TAG: &[u8] = b"litchi-ppt/scoped/block-root/v1";

fn sha256(bytes: &[u8]) -> Digest {
    Sha256::digest(bytes).into()
}

/// Block `index` of `bytes`: 512 bytes, or the shorter tail.
fn block(bytes: &[u8], index: usize) -> &[u8] {
    let start = index.saturating_mul(BLOCK).min(bytes.len());
    let stop = start.saturating_add(BLOCK).min(bytes.len());
    bytes.get(start..stop).unwrap_or_default()
}

/// The SHA-256 of every 512-byte block of one exact image, with its length
/// and its root, computed once.
pub(crate) struct BlockDigests {
    len: usize,
    blocks: Box<[Digest]>,
    root: OnceLock<Digest>,
}

impl BlockDigests {
    /// Hash every block of `bytes`.
    pub(crate) fn of(bytes: &[u8]) -> Self {
        let blocks = bytes.chunks(BLOCK).map(sha256).collect();
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
    pub(crate) fn derive(&self, bytes: &[u8], changed: &[Range<usize>]) -> Self {
        let count = bytes.len().div_ceil(BLOCK);
        let mut blocks = Vec::with_capacity(count);
        let mut spans = changed.iter().peekable();
        for index in 0..count {
            while spans.next_if(|span| span.end <= index).is_some() {}
            let is_changed = spans.peek().is_some_and(|span| span.start <= index);
            let data = block(bytes, index);
            let reused = (!is_changed && self.extent(index) == data.len())
                .then(|| self.blocks.get(index))
                .flatten();
            blocks.push(reused.copied().unwrap_or_else(|| sha256(data)));
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
    changed: Arc<[Range<usize>]>,
}

/// The union of two sorted, disjoint block-span lists, merged.
fn union(left: &[Range<usize>], right: &[Range<usize>]) -> Arc<[Range<usize>]> {
    let mut all: Vec<Range<usize>> = left.iter().chain(right).cloned().collect();
    all.sort_unstable_by_key(|span| span.start);
    let mut merged: Vec<Range<usize>> = Vec::with_capacity(all.len());
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

    /// The recipe of an image that equals `bytes` (this package's bytes)
    /// outside the block spans `changed`. It reuses this package's list
    /// when known, or composes with this package's own recipe (the union of
    /// both span lists), so no block is hashed here.
    pub(crate) fn child(&self, bytes: &[u8], changed: &[Range<usize>]) -> Recipe {
        if self.blocks.get().is_none()
            && let Some(recipe) = self.recipe.get()
        {
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
    use super::{BLOCK, BlockDigests, ExactCache, Recipe, union};
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
    fn a_derived_list_equals_the_full_list() {
        let before = image(6 * BLOCK + 100, 3);
        let source = BlockDigests::of(&before);
        let mut after = before.clone();
        after[2 * BLOCK + 1] ^= 1;
        let derived = source.derive(&after, &[2..3]);
        assert_eq!(derived.root(), BlockDigests::of(&after).root());
        // Growth past a partial tail, shrinkage, and the empty image.
        let mut grown = before.clone();
        grown.extend_from_slice(&image(2 * BLOCK, 9));
        assert_eq!(
            source.derive(&grown, &[6..9]).root(),
            BlockDigests::of(&grown).root()
        );
        let shrunk = &before[..2 * BLOCK + 3];
        assert_eq!(
            source.derive(shrunk, &[2..7]).root(),
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
                changed: Arc::from([1..2]),
            },
        );
        assert_eq!(cache.root(&second), root);
        let child = cache.child(&second, &[4..5]);
        assert_eq!(&*child.changed, &[1..2, 4..5]);
        let derived = ExactCache::default();
        derived.seed_derived(BlockDigests::of(&third).root(), child);
        assert_eq!(
            derived.blocks(&third).root(),
            BlockDigests::of(&third).root()
        );
        assert_eq!(&*union(&[0..2, 5..6], &[1..3, 6..7]), &[0..3, 5..7]);
    }
}
