//! The alternative-format payload family of a source-backed package
//! (record 0355a): replace the opaque payload behind one existing
//! main-story `w:altChunk`, keeping its relationship, part name and media
//! type.
//!
//! [`Package::edit_alt_chunk_payload`](crate::source_backed::Package::edit_alt_chunk_payload)
//! stages an [`Edit`] whose commit returns a new source-backed package;
//! writing it rewrites only that part. The payload is inert: it is never
//! imported, parsed, rendered, fetched or executed.

use std::sync::Arc;

use litchi_opc::PackURI;

use crate::Result;
use crate::alt::Data;
use crate::source_backed::deferred::{Change, source_family};
use crate::source_backed::{AltChunkSelector, Package};

/// One payload part's exact before and after bytes.
#[derive(Clone)]
pub(crate) struct PayloadChange {
    part: PackURI,
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
}

impl Change for PayloadChange {
    type Error = crate::Error;

    fn is_noop(&self) -> bool {
        self.before.as_slice() == self.after.as_slice()
    }

    fn inverse(&self) -> Self {
        Self {
            part: self.part.clone(),
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }

    fn publish(&self, package: &Package) -> Result<Package> {
        if package.part_payload(&self.part)?.as_slice() != self.before.as_slice() {
            return Err(crate::Error::PatchConflict {
                operation: "apply_alt_chunk_payload",
                reason: "the payload part does not hold the patch's before bytes",
            });
        }
        if self.is_noop() {
            return Ok(package.clone());
        }
        let (part, after) = (self.part.clone(), Arc::clone(&self.after));
        package.deferred_part_overlay(part, after)
    }
}

source_family!("alternative-format payload", PayloadChange);

/// The budgets of this family, reported by
/// [`crate::Error::AltChunkLimit`].
pub use crate::alt::LimitKind;

/// A staged replacement of one altChunk payload. The source is never
/// changed: [`Self::commit`] returns a new package.
#[derive(Clone)]
pub struct Edit {
    package: Package,
    selector: AltChunkSelector,
    /// The staged payload with its media type and extension.
    replacement: Option<(Arc<Vec<u8>>, &'static str, &'static str)>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    /// The package the edit is staged over.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Stage `replacement` as the anchor's payload.
    ///
    /// # Errors
    ///
    /// Returns an error for a payload that is not valid for its media type.
    pub fn replace(&mut self, replacement: Data) -> Result<&mut Self> {
        replacement.validate()?;
        let (media_type, extension) = (replacement.media_type(), replacement.extension());
        self.replacement = Some((Arc::new(replacement.into_bytes()), media_type, extension));
        Ok(self)
    }

    /// Discard the staged replacement.
    pub fn reset(&mut self) {
        self.replacement = None;
    }

    /// Commit the replacement as a new source-backed package; nothing staged
    /// commits the source unchanged.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, relationship, target-layout, media-type,
    /// signature or source error before any output.
    pub fn commit(self) -> Result<Commit> {
        let Some((after, media_type, extension)) = self.replacement else {
            return Ok(Commit::new(
                self.package.clone(),
                Patch {
                    change: PayloadChange {
                        part: self.package.main_part_name()?,
                        before: Arc::new(Vec::new()),
                        after: Arc::new(Vec::new()),
                    },
                    restore: Some((self.package.clone(), self.package)),
                },
                Diagnostics::new(false),
            ));
        };
        let part = self
            .package
            .alt_chunk_payload_target(self.selector, media_type, extension)?;
        let before = self.package.part_payload(&part)?;
        commit_change(
            &self.package,
            PayloadChange {
                part,
                before,
                after,
            },
        )
    }
}

impl Package {
    /// Stage a replacement of the payload behind one active main-story
    /// altChunk anchor.
    #[must_use]
    pub fn edit_alt_chunk_payload(&self, selector: impl Into<AltChunkSelector>) -> Edit {
        Edit {
            package: self.clone(),
            selector: selector.into(),
            replacement: None,
        }
    }

    /// Replay a payload patch onto this package, which must hold the
    /// patch's before bytes.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for another payload.
    pub fn apply_alt_chunk_payload(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}
