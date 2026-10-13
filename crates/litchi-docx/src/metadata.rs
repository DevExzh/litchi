//! Core document properties (`docProps/core.xml`) as a scoped transaction.
//!
//! [`Package::metadata`](crate::Package::metadata) reads the typed core
//! properties, keeping a package without a core-properties part distinct from
//! one with an empty part. Every change goes through the ADR 0034 §7 family:
//! [`Package::edit_metadata`](crate::Package::edit_metadata) stages an
//! [`Edit`], [`Edit::commit`] publishes a new [`Package`] and
//! leaves the source untouched, and
//! [`Package::apply_metadata`](crate::Package::apply_metadata) replays a
//! [`Patch`] onto another snapshot with the same base value.
//!
//! The family is a thin DOCX wrapper over the shared OOXML codec
//! ([`litchi_ooxml_common::properties`]): the codec keeps the part's dialect,
//! name and relationship identity, refuses shared or ambiguous graphs, and
//! publishes nothing for a value equal to the stored one.
//!
//! ```
//! use litchi_docx::Package;
//!
//! let package = Package::new()?;
//! let mut edit = package.edit_metadata();
//! edit.update(|props| props.title = Some("Quarterly report".to_owned()));
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! let next = commit.into_package();
//! assert_eq!(
//!     next.metadata().and_then(|props| props.title.as_deref()),
//!     Some("Quarterly report")
//! );
//! // The source snapshot is unchanged.
//! assert_eq!(package.metadata().and_then(|props| props.title.as_deref()), None);
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub use litchi_ooxml_common::properties::Props;

use crate::Package;

/// The largest accepted text value of one core property, in UTF-8 bytes.
///
/// This is the shared OOXML codec's per-property bound; the transaction
/// checks it before it stages any package change so an oversized value is a
/// typed [`Error::Limit`] instead of a codec failure.
pub const MAX_PROPERTY_TEXT_BYTES: usize = 1_048_576;

/// Result alias of the metadata family.
pub type Result<T> = std::result::Result<T, Error>;

/// The bounded resource a metadata transaction reports.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// UTF-8 bytes of one text-valued core property.
    PropertyTextBytes,
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::PropertyTextBytes => "property text bytes",
        })
    }
}

/// A metadata transaction failure.
///
/// Messages carry structural metadata only, never property values.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The package refused the change or its core-properties graph is
    /// invalid.
    #[error(transparent)]
    Package(#[from] crate::Error),
    /// A changed write was attempted on a signed package. Remove the
    /// signature explicitly with [`Package::unsign`](crate::Package::unsign)
    /// first; an exact no-op keeps the signature.
    #[error("changing the core properties of a signed DOCX package requires Package::unsign")]
    Signed,
    /// The patch's base value does not match the package it is applied to.
    #[error("the metadata patch base does not match the package's core properties")]
    PatchConflict,
    /// A staged value exceeds a metadata bound.
    #[error("metadata {kind} limit exceeded: {observed} > {maximum}")]
    Limit {
        /// Bounded resource.
        kind: LimitKind,
        /// Observed value.
        observed: usize,
        /// Maximum accepted value.
        maximum: usize,
    },
}

/// What a metadata commit changed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    /// Whether the commit published a different core-properties value.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }
}

/// A replayable core-properties change: the value the edit started from and
/// the value it published. `None` is an absent core-properties part.
#[derive(Debug, Clone, PartialEq)]
pub struct Patch {
    base: Option<Props>,
    target: Option<Props>,
}

impl Patch {
    /// The core properties the patch expects to find.
    #[must_use]
    pub fn base(&self) -> Option<&Props> {
        self.base.as_ref()
    }

    /// The core properties the patch publishes.
    #[must_use]
    pub fn target(&self) -> Option<&Props> {
        self.target.as_ref()
    }

    /// Whether applying the patch leaves every package byte unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.base == self.target
    }

    /// The patch that restores the base value.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target.clone(),
            target: self.base.clone(),
        }
    }
}

/// A published metadata change.
#[derive(Clone)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// The package with the staged core properties.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// The replayable change.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// What the commit changed.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// The published package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// The published package and its patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

impl std::fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}

/// A staged core-properties change over one package snapshot.
///
/// Dropping the edit discards it; the source package is never mutated.
#[derive(Clone)]
pub struct Edit {
    source: Package,
    base: Option<Props>,
    staged: Option<Props>,
}

impl Edit {
    pub(crate) fn new(source: Package, base: Option<Props>) -> Self {
        Self {
            staged: base.clone(),
            source,
            base,
        }
    }

    /// The staged core properties.
    #[must_use]
    pub fn metadata(&self) -> Option<&Props> {
        self.staged.as_ref()
    }

    /// Replace the staged core properties with `props`.
    pub fn set(&mut self, props: Props) -> &mut Self {
        self.staged = Some(props);
        self
    }

    /// Change the staged core properties in place, starting from an empty
    /// value when the package has none.
    pub fn update(&mut self, change: impl FnOnce(&mut Props)) -> &mut Self {
        change(self.staged.get_or_insert_with(Props::default));
        self
    }

    /// Stage the removal of the core-properties part.
    pub fn clear(&mut self) -> &mut Self {
        self.staged = None;
        self
    }

    /// Publish the staged value as a new package.
    ///
    /// An unchanged value returns a package that shares every part with the
    /// source, signature included.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] for an oversized text property,
    /// [`Error::Signed`] for a change to a signed package, and
    /// [`Error::Package`] when the codec or package graph refuses the value.
    pub fn commit(self) -> Result<Commit> {
        let patch = Patch {
            base: self.base,
            target: self.staged,
        };
        let (package, changed) = self.source.publish_metadata(patch.target.as_ref())?;
        Ok(Commit {
            package,
            patch,
            diagnostics: Diagnostics { changed },
        })
    }
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("present", &self.staged.is_some())
            .finish_non_exhaustive()
    }
}

/// Check every text-valued property against [`MAX_PROPERTY_TEXT_BYTES`].
pub(crate) fn admit(props: &Props) -> Result<()> {
    let texts = [
        props.title.as_deref(),
        props.subject.as_deref(),
        props.creator.as_deref(),
        props.description.as_deref(),
        props.identifier.as_deref(),
        props.last_modified_by.as_deref(),
        props.category.as_deref(),
        props.content_status.as_deref(),
        props.revision.as_deref(),
        props.version.as_deref(),
        props.language.as_deref(),
    ];
    let keywords = props
        .keywords
        .as_ref()
        .map_or(0, |keywords| keywords.joined().len());
    let observed = texts
        .into_iter()
        .flatten()
        .map(str::len)
        .chain(std::iter::once(keywords))
        .max()
        .unwrap_or(0);
    if observed > MAX_PROPERTY_TEXT_BYTES {
        return Err(Error::Limit {
            kind: LimitKind::PropertyTextBytes,
            observed,
            maximum: MAX_PROPERTY_TEXT_BYTES,
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn edit_commit_publishes_a_new_package_and_keeps_the_source() -> Result<()> {
        let package = Package::new()?;
        let before = package.metadata().cloned();
        let mut edit = package.edit_metadata();
        edit.update(|props| {
            props.title = Some("T".to_owned());
            props.creator = Some("A".to_owned());
        });
        let commit = edit.commit()?;
        assert!(commit.diagnostics().changed());
        assert!(!commit.patch().is_noop());
        assert_eq!(package.metadata().cloned(), before);
        let mut next = commit.package().clone();
        assert_eq!(
            next.metadata().and_then(|props| props.title.as_deref()),
            Some("T")
        );

        let bytes = next.to_bytes()?;
        let reopened = Package::from_bytes(&bytes)?;
        assert_eq!(
            reopened
                .metadata()
                .and_then(|props| props.creator.as_deref()),
            Some("A")
        );
        Ok(())
    }

    #[test]
    fn unchanged_commit_is_an_exact_noop() -> Result<()> {
        let package = Package::new()?;
        let commit = package.edit_metadata().commit()?;
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert_eq!(
            package.clone().to_bytes()?,
            commit.package().clone().to_bytes()?
        );
        Ok(())
    }

    #[test]
    fn clear_removes_the_part_and_inverse_restores_it() -> Result<()> {
        let package = Package::new()?;
        let mut edit = package.edit_metadata();
        edit.clear();
        let (cleared, patch) = edit.commit()?.into_parts();
        assert!(cleared.metadata().is_none());
        let restored = cleared.apply_metadata(&patch.inverse())?;
        assert_eq!(restored.metadata(), package.metadata());
        Ok(())
    }

    #[test]
    fn apply_refuses_a_conflicting_base() -> Result<()> {
        let package = Package::new()?;
        let mut first = package.edit_metadata();
        first.update(|props| props.title = Some("one".to_owned()));
        let (one, _) = first.commit()?.into_parts();
        let mut second = package.edit_metadata();
        second.update(|props| props.title = Some("two".to_owned()));
        let patch = second.commit()?.patch().clone();
        assert!(matches!(
            one.apply_metadata(&patch),
            Err(Error::PatchConflict)
        ));
        let replayed = package.apply_metadata(&patch)?;
        assert_eq!(
            replayed.metadata().and_then(|props| props.title.as_deref()),
            Some("two")
        );
        Ok(())
    }

    #[test]
    fn oversized_text_is_a_typed_limit() -> Result<()> {
        let package = Package::new()?;
        let mut edit = package.edit_metadata();
        edit.update(|props| props.title = Some("x".repeat(MAX_PROPERTY_TEXT_BYTES + 1)));
        assert!(matches!(
            edit.commit(),
            Err(Error::Limit {
                kind: LimitKind::PropertyTextBytes,
                ..
            })
        ));
        Ok(())
    }
}
