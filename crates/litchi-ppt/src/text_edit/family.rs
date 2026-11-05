//! The source-backed text family (record 0357a): a same-length splice of one
//! shape's text atom, published as a new [`Package`] over the composed
//! overlay source without materializing the package.
//!
//! [`Package::edit_text`] opens the family. Its commit plans one
//! [`litchi_cfb::ValidatedOverlayPlan`] splice in the `PowerPoint Document`
//! stream, reopens the composed target and reads the selected text back, and
//! then adopts that validated compound-file index as the published package:
//! the published package reads its bytes positionally through the composed
//! source, and its `write_to`/`save` stream them from there.
//!
//! The patch is record-sized: the encoded text atom's offset and its two
//! byte images, bound to the exact source and target artifact fingerprints
//! (SHA-256 of the complete compound file). It applies to any package with
//! the exact source fingerprint.

use std::fmt;

use litchi_cfb::ArtifactFingerprint;

use super::{
    Error, Result, SourceBackedCommit, SourceBackedOptions, SourceBackedPatch, SourceSnapshot,
    SourceTransaction, Target, publish_source_operation,
};
use crate::Package;
use crate::package::{Error as PackageError, OpenOptions, SourceBackedPackage};

/// A text-family limit dimension.
///
/// The splice is bounded by the default [`litchi_cfb::StreamSpliceLimits`]
/// and the republished package by the source package's
/// [`SourceLimits`](crate::SourceLimits), so no caller-selected dimension
/// exists.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {}

/// Content-free diagnostics of one source-backed text commit.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    target: Target,
    changed_spans: usize,
    replacement_bytes: u64,
    source_bytes: u64,
    source_fingerprint: ArtifactFingerprint,
    target_fingerprint: ArtifactFingerprint,
}

impl Diagnostics {
    fn from_native(commit: &SourceBackedCommit) -> Self {
        let native = commit.diagnostics();
        Self {
            changed: !commit.is_noop(),
            target: native.target(),
            changed_spans: native.changed_spans(),
            replacement_bytes: native.replacement_bytes(),
            source_bytes: native.source_bytes(),
            source_fingerprint: native.source_fingerprint(),
            target_fingerprint: native.target_fingerprint(),
        }
    }

    fn noop(target: Target, source_bytes: u64, fingerprint: ArtifactFingerprint) -> Self {
        Self {
            changed: false,
            target,
            changed_spans: 0,
            replacement_bytes: 0,
            source_bytes,
            source_fingerprint: fingerprint,
            target_fingerprint: fingerprint,
        }
    }

    /// Whether the package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// The edited shape.
    #[must_use]
    pub const fn target(self) -> Target {
        self.target
    }

    /// Number of physical compound-file spans the splice changed.
    #[must_use]
    pub const fn changed_spans(self) -> usize {
        self.changed_spans
    }

    /// Encoded replacement bytes submitted to the splice planner.
    #[must_use]
    pub const fn replacement_bytes(self) -> u64 {
        self.replacement_bytes
    }

    /// Length of the complete source compound file (the target has the
    /// same length).
    #[must_use]
    pub const fn source_bytes(self) -> u64 {
        self.source_bytes
    }

    /// Fingerprint of the exact source artifact.
    #[must_use]
    pub const fn source_fingerprint(self) -> ArtifactFingerprint {
        self.source_fingerprint
    }

    /// Fingerprint of the exact published artifact.
    #[must_use]
    pub const fn target_fingerprint(self) -> ArtifactFingerprint {
        self.target_fingerprint
    }
}

/// An exact, reversible same-length text splice, bound to the exact source
/// artifact fingerprint it was made from.
#[derive(Clone)]
pub struct Patch {
    shape: Target,
    source: ArtifactFingerprint,
    target: ArtifactFingerprint,
    operation: Option<SourceBackedPatch>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("source_fingerprint", &self.source)
            .field("target_fingerprint", &self.target)
            .field("is_noop", &self.is_noop())
            .finish()
    }
}

impl Patch {
    fn from_native(shape: Target, patch: SourceBackedPatch) -> Self {
        Self {
            shape,
            source: patch.source_fingerprint(),
            target: patch.target_fingerprint(),
            operation: (!patch.is_empty()).then_some(patch),
        }
    }

    /// Whether applying the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.operation.is_none()
    }

    /// The patch that undoes this one, applicable to its result.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            shape: self.shape,
            source: self.target,
            target: self.source,
            operation: self.operation.as_ref().map(SourceBackedPatch::inverse),
        }
    }

    /// The edited shape.
    #[must_use]
    pub const fn target(&self) -> Target {
        self.shape
    }

    /// Fingerprint of the exact source artifact the patch applies to.
    #[must_use]
    pub const fn source_fingerprint(&self) -> ArtifactFingerprint {
        self.source
    }

    /// Fingerprint of the exact artifact the patch produces.
    #[must_use]
    pub const fn target_fingerprint(&self) -> ArtifactFingerprint {
        self.target
    }
}

/// A published text edit: the new package, its exact patch and diagnostics.
#[derive(Clone)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("package", &self.package)
            .field("patch", &self.patch)
            .field("diagnostics", &self.diagnostics)
            .finish()
    }
}

impl Commit {
    /// The published package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// The exact patch from the source package to the published one.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Consume the commit into its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit into its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

/// A staged same-length replacement of one shape's text.
///
/// Dropping the edit discards it.
pub struct Edit {
    package: Package,
    state: State,
}

enum State {
    /// An unsigned package: the source-backed splice owner.
    Source(SourceTransaction),
    /// A signed package: only an exact no-op can be published.
    Signed {
        target: Target,
        text: String,
        replacement: Option<String>,
        fingerprint: ArtifactFingerprint,
    },
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("package", &self.package)
            .field("target", &self.target())
            .field("signed", &matches!(self.state, State::Signed { .. }))
            .finish_non_exhaustive()
    }
}

impl Edit {
    /// The text currently held by the selected atom.
    #[must_use]
    pub fn text(&self) -> &str {
        match &self.state {
            State::Source(transaction) => transaction.text(),
            State::Signed { text, .. } => text,
        }
    }

    /// The selected shape.
    #[must_use]
    pub const fn target(&self) -> Target {
        match &self.state {
            State::Source(transaction) => transaction.target(),
            State::Signed { target, .. } => *target,
        }
    }

    /// Stage a replacement whose encoded atom length is unchanged.
    ///
    /// # Errors
    ///
    /// [`super::Refusal::LengthChange`] for a replacement of another encoded length
    /// and [`super::Refusal::IncompatibleEncoding`] for one the atom's encoding
    /// cannot hold; a failed call leaves the edit unchanged.
    pub fn set_text(&mut self, value: impl Into<String>) -> Result<()> {
        match &mut self.state {
            State::Source(transaction) => transaction.set_text(value),
            State::Signed { replacement, .. } => {
                *replacement = Some(value.into());
                Ok(())
            },
        }
    }

    /// Publish the edit as a new package.
    ///
    /// An exact no-op returns the source package unchanged. A change is
    /// planned as one overlay splice, the composed target is reopened and
    /// read back, and the published package reads through it.
    ///
    /// # Errors
    ///
    /// [`Error::Signed`] for a changed commit on a signed package, and the
    /// splice planner's or the republished package's error otherwise.
    pub fn commit(self) -> Result<Commit> {
        match self.state {
            State::Source(transaction) => publish(&self.package, transaction.commit()?),
            State::Signed {
                target,
                text,
                replacement,
                fingerprint,
            } => {
                if replacement.is_some_and(|value| value != text) {
                    return Err(Error::Signed);
                }
                Ok(noop_commit(&self.package, target, fingerprint))
            },
        }
    }
}

impl Package {
    /// Start a same-length text splice of one shape's text atom
    /// (record 0357a).
    ///
    /// The edit resolves the atom through bounded positional reads of the
    /// retained compound file; it never copies the package. A signed
    /// package opens for exact no-op commits.
    ///
    /// # Errors
    ///
    /// [`Error::Package`] for an encrypted package, a typed [`super::Refusal`] for
    /// an absent or unsupported target or an unsupported source (macros,
    /// embedded storages), and [`Error::Source`] for a changed source.
    pub fn edit_text(&self, target: Target) -> Result<Edit> {
        if self.is_encrypted() {
            return Err(Error::Package(PackageError::EncryptionPolicy));
        }
        let state = if is_signed(self) {
            let bytes = self.to_encrypted_bytes()?;
            let text = super::inspect_shape_text_readonly(&bytes, target, open_options(self))?;
            State::Signed {
                target,
                text,
                replacement: None,
                fingerprint: signed_fingerprint(self)?,
            }
        } else {
            State::Source(source_snapshot(self)?.edit_text(target)?)
        };
        Ok(Edit {
            package: self.clone(),
            state,
        })
    }

    /// Replay a text patch on this package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] unless this package's exact artifact
    /// fingerprint is the patch source, and [`Error::Signed`] for a changed
    /// patch on a signed package.
    pub fn apply_text(&self, patch: &Patch) -> Result<Commit> {
        if self.is_encrypted() {
            return Err(Error::Package(PackageError::EncryptionPolicy));
        }
        if is_signed(self) {
            if signed_fingerprint(self)? != patch.source {
                return Err(Error::PatchConflict);
            }
            if !patch.is_noop() {
                return Err(Error::Signed);
            }
            return Ok(noop_commit(self, patch.shape, patch.source).with_patch(patch));
        }
        let current = source_snapshot(self)?;
        if current.fingerprint() != patch.source {
            return Err(Error::PatchConflict);
        }
        let Some(operation) = &patch.operation else {
            return Ok(noop_commit(self, patch.shape, patch.source).with_patch(patch));
        };
        let native = publish_source_operation(current, operation.operation_clone()?)?;
        if native.snapshot().fingerprint() != patch.target {
            return Err(Error::PatchConflict);
        }
        publish(self, native)
    }
}

impl Commit {
    fn with_patch(mut self, patch: &Patch) -> Self {
        self.patch = patch.clone();
        self
    }
}

impl SourceBackedPatch {
    fn operation_clone(&self) -> Result<super::SourceOperation> {
        self.operation.clone().ok_or_else(|| {
            Error::Package(PackageError::Corrupted(
                "PPT text patch has no splice operation".into(),
            ))
        })
    }
}

fn publish(source: &Package, native: SourceBackedCommit) -> Result<Commit> {
    let diagnostics = Diagnostics::from_native(&native);
    let package = if native.is_noop() {
        source.clone()
    } else {
        let options = open_options(source);
        let shared = std::sync::Arc::clone(&native.snapshot().inner.shared);
        Package::from_source(
            SourceBackedPackage::from_validated_shared(shared, options)?,
            options,
        )?
    };
    let patch = Patch::from_native(diagnostics.target, native.patch().clone());
    Ok(Commit {
        package,
        patch,
        diagnostics,
    })
}

fn noop_commit(package: &Package, target: Target, fingerprint: ArtifactFingerprint) -> Commit {
    Commit {
        package: package.clone(),
        patch: Patch {
            shape: target,
            source: fingerprint,
            target: fingerprint,
            operation: None,
        },
        diagnostics: Diagnostics::noop(target, package.shared_ole_file().file_size(), fingerprint),
    }
}

fn open_options(package: &Package) -> OpenOptions<'static> {
    let options = OpenOptions::new()
        .source_limits(package.source_limits())
        .leniency(package.leniency());
    #[cfg(feature = "vba-inspection")]
    let options = options.vba_project_limits(package.vba_project_limits());
    options
}

fn source_snapshot(package: &Package) -> Result<SourceSnapshot> {
    SourceSnapshot::from_shared(
        package.shared_ole_file_arc(),
        SourceBackedOptions {
            source_limits: package.source_limits(),
            ..SourceBackedOptions::default()
        },
    )
}

fn is_signed(package: &Package) -> bool {
    package
        .shared_ole_file()
        .directory_entries()
        .any(|entry| crate::package::codec::is_signature_component(&entry.name))
}

/// The complete-artifact fingerprint of a signed package, read twice and
/// compared (the overlay publisher, which fingerprints unsigned sources,
/// refuses a signed one).
fn signed_fingerprint(package: &Package) -> Result<ArtifactFingerprint> {
    let shared = package.shared_ole_file();
    let first = shared
        .caller_bracketed_identity_fingerprint()
        .map_err(Error::Source)?;
    let second = shared
        .caller_bracketed_identity_fingerprint()
        .map_err(Error::Source)?;
    if first != second {
        return Err(Error::Source(litchi_cfb::OverlayError::Unavailable {
            reason: "PPT source changed while it was fingerprinted".into(),
        }));
    }
    Ok(first)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests;
