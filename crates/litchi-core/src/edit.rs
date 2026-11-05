//! Format-neutral edit vocabulary: commit triad, diagnostics and refusals.
//!
//! Every format publishes a mutation the same way (ADR 0034 §3, record 0351):
//! an owner snapshot `Package` stages an `Edit`, `Edit::commit()` returns a
//! `Commit` that carries the **new** `Package`, the reversible `Patch` and
//! content-free `Diagnostics`, and the source `Package` is never mutated.
//! Dropping an `Edit` discards it; there is no `rollback`.
//!
//! This module holds the parts of that contract that do not depend on any
//! format:
//!
//! - [`Commit`], a generic commit triad a scoped family may re-export as its
//!   own `Commit` through a type alias
//!   (`pub type Commit = litchi_core::edit::Commit<Package, Patch>;`) when it
//!   carries nothing beyond the triad;
//! - [`Diagnostics`], the minimal content-free commit diagnostics (whether the
//!   edit changed its source), for families that report nothing else;
//! - [`Refusal`] and [`RefusalKind`], the typed reason an edit route refuses a
//!   well-formed request. They replace the per-crate `UnsafeEdit` variants:
//!   a module error carries a refusal as `Error::Refused(Refusal)`, beside
//!   the shared `PatchConflict`, `SourceChanged`, `Limit` and `Signed`
//!   variants, which keep their own names because they mean other things.
//!
//! A family with richer diagnostics or a commit that carries more than the
//! triad keeps its own types; it still exposes the same accessors
//! (`package`, `into_package`, `patch`, `diagnostics`, `into_parts`, and
//! `Diagnostics::changed`).
//!
//! # Content redaction
//!
//! A [`Refusal`] holds only `&'static str` text and a [`RefusalKind`], so it
//! cannot carry authored content (user text, entry names, sheet names) by
//! construction. [`Diagnostics`] holds only a flag.
//!
//! ```
//! use litchi_core::edit::{Commit, Diagnostics, Refusal, RefusalKind};
//!
//! // A family whose commit is exactly the triad aliases the shared type.
//! type NotesCommit = Commit<String, Vec<u8>>;
//! let commit = NotesCommit::new("new package".to_owned(), vec![1], Diagnostics::new(true));
//! assert!(commit.diagnostics().changed());
//! let (package, patch) = commit.into_parts();
//! assert_eq!((package.as_str(), patch.as_slice()), ("new package", &[1][..]));
//!
//! // A refusal is typed for matching and content-free for display.
//! let refusal = Refusal::new(
//!     RefusalKind::UnsupportedContent,
//!     "set_shape_text",
//!     "the shape text holds a field",
//! );
//! assert_eq!(refusal.kind(), RefusalKind::UnsupportedContent);
//! assert_eq!(
//!     refusal.to_string(),
//!     "set_shape_text refused (unsupported_content): the shape text holds a field",
//! );
//! ```

use std::fmt;

/// Content-free diagnostics of one commit or patch application.
///
/// This is the minimal form of the `Diagnostics` member of the commit triad:
/// it reports only whether the edit changed its source. An exact no-op
/// commit reports `changed() == false`, and its package is the source
/// package unchanged (byte-identical output, signatures preserved).
///
/// A family that reports more (touched parts, counts, repairs) defines its
/// own `Diagnostics` type with the same `changed()` accessor.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    /// Diagnostics of an exact no-op.
    pub const UNCHANGED: Self = Self { changed: false };

    /// Diagnostics of an edit that changed its source.
    pub const CHANGED: Self = Self { changed: true };

    /// Diagnostics reporting whether the edit changed its source.
    #[must_use]
    pub const fn new(changed: bool) -> Self {
        Self { changed }
    }

    /// Whether the edit changed its source. `false` is an exact no-op.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }
}

/// The commit triad: the new package, its reversible patch and diagnostics.
///
/// `P` is the format's `Package`, `T` the family's `Patch` and `D` its
/// `Diagnostics`. The source package is never part of a commit: it is
/// unchanged, and the caller still holds it.
///
/// A commit is plain data. Constructing one does not apply anything, and a
/// patch is checked against its base when it is applied, so a commit
/// assembled outside its format cannot publish an unchecked change.
/// Debug output reports only diagnostics; it never formats the retained
/// package or patch payloads.
#[derive(Clone, PartialEq, Eq)]
#[must_use = "a commit holds the new package; dropping it discards the edit"]
pub struct Commit<P, T, D = Diagnostics> {
    package: P,
    patch: T,
    diagnostics: D,
}

impl<P, T, D: fmt::Debug> fmt::Debug for Commit<P, T, D> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        // Owners and patches may retain authored payloads. Only the family's
        // content-free diagnostics belong in a commit's diagnostic rendering.
        formatter
            .debug_struct("Commit")
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}

impl<P, T, D> Commit<P, T, D> {
    /// Assemble a commit from its triad.
    pub const fn new(package: P, patch: T, diagnostics: D) -> Self {
        Self {
            package,
            patch,
            diagnostics,
        }
    }

    /// The new package.
    #[must_use]
    pub const fn package(&self) -> &P {
        &self.package
    }

    /// The reversible patch from the source package to [`Self::package`].
    #[must_use]
    pub const fn patch(&self) -> &T {
        &self.patch
    }

    /// The commit's content-free diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &D {
        &self.diagnostics
    }

    /// The new package, discarding the patch and diagnostics.
    #[must_use]
    pub fn into_package(self) -> P {
        self.package
    }

    /// The new package and the patch, discarding the diagnostics.
    #[must_use]
    pub fn into_parts(self) -> (P, T) {
        (self.package, self.patch)
    }
}

/// Why an edit route refuses a well-formed request.
///
/// The kind is the stable, matchable part of a [`Refusal`]; the litchi-ops
/// layer classifies each kind into exactly one unified error code, so every
/// format reports the same refusal the same way:
///
/// | Kind | litchi-ops code |
/// |---|---|
/// | [`UnsupportedOperation`](Self::UnsupportedOperation) | `unsupported_capability` |
/// | [`UnsupportedContent`](Self::UnsupportedContent) | `unsupported_capability` |
/// | [`UnsupportedTarget`](Self::UnsupportedTarget) | `unsupported_type` |
/// | [`UnsupportedProperty`](Self::UnsupportedProperty) | `unsupported_property` |
/// | [`InvalidValue`](Self::InvalidValue) | `invalid_value` |
/// | [`NotFound`](Self::NotFound) | `not_found` |
/// | [`Protected`](Self::Protected) | `document_protected` |
///
/// Conditions that are not refusals keep their own variants of the module
/// error: a patch whose base is not this package is `PatchConflict`, a source
/// file that changed after open is `SourceChanged`, an exceeded budget is
/// `Limit`, and a changed write on a signed package is `Signed`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum RefusalKind {
    /// The route does not implement this operation, or not in this context
    /// (for example a source-backed edit that admits one operation per
    /// commit).
    UnsupportedOperation,
    /// The source holds content the route cannot carry through losslessly
    /// (markup-compatibility branches, unknown children, opaque records).
    UnsupportedContent,
    /// The target exists but is not of a kind that takes the operation (a
    /// shape without a text body, a slide without a notes page).
    UnsupportedTarget,
    /// The property cannot be set or read through this route.
    UnsupportedProperty,
    /// The value is outside what the format can store.
    InvalidValue,
    /// The selected target does not exist.
    NotFound,
    /// Document, sheet or write protection is enforced.
    Protected,
}

impl RefusalKind {
    /// Every kind, in declaration order.
    pub const ALL: &'static [Self] = &[
        Self::UnsupportedOperation,
        Self::UnsupportedContent,
        Self::UnsupportedTarget,
        Self::UnsupportedProperty,
        Self::InvalidValue,
        Self::NotFound,
        Self::Protected,
    ];

    /// The stable `snake_case` name of the kind.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::UnsupportedOperation => "unsupported_operation",
            Self::UnsupportedContent => "unsupported_content",
            Self::UnsupportedTarget => "unsupported_target",
            Self::UnsupportedProperty => "unsupported_property",
            Self::InvalidValue => "invalid_value",
            Self::NotFound => "not_found",
            Self::Protected => "protected",
        }
    }
}

impl fmt::Display for RefusalKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// A typed, content-free edit refusal.
///
/// Carried by a module error as `Error::Refused(Refusal)`. The operation and
/// reason are static text written by the format, never caller content.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Refusal {
    kind: RefusalKind,
    operation: &'static str,
    reason: &'static str,
}

impl Refusal {
    /// A refusal of `operation` for `reason`.
    #[must_use]
    pub const fn new(kind: RefusalKind, operation: &'static str, reason: &'static str) -> Self {
        Self {
            kind,
            operation,
            reason,
        }
    }

    /// The refusal's kind, for matching and classification.
    #[must_use]
    pub const fn kind(&self) -> RefusalKind {
        self.kind
    }

    /// The refused operation, such as `set_shape_text`.
    #[must_use]
    pub const fn operation(&self) -> &'static str {
        self.operation
    }

    /// Why the route refuses it.
    #[must_use]
    pub const fn reason(&self) -> &'static str {
        self.reason
    }
}

impl fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "{} refused ({}): {}",
            self.operation, self.kind, self.reason
        )
    }
}

impl std::error::Error for Refusal {}

#[cfg(test)]
mod tests {
    use std::fmt;

    use super::{Commit, Diagnostics, Refusal, RefusalKind};

    #[test]
    fn diagnostics_report_only_the_change_flag() {
        assert!(!Diagnostics::default().changed());
        assert_eq!(Diagnostics::default(), Diagnostics::UNCHANGED);
        assert!(Diagnostics::CHANGED.changed());
        assert_eq!(Diagnostics::new(true), Diagnostics::CHANGED);
        assert_eq!(
            format!("{:?}", Diagnostics::CHANGED),
            "Diagnostics { changed: true }"
        );
    }

    #[test]
    fn commit_exposes_the_triad_and_splits() {
        let commit = Commit::new(7_u32, "patch", Diagnostics::CHANGED);
        assert_eq!(*commit.package(), 7);
        assert_eq!(*commit.patch(), "patch");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.clone().into_package(), 7);
        assert_eq!(commit.into_parts(), (7, "patch"));
    }

    #[test]
    fn commit_accepts_family_diagnostics() {
        #[derive(Debug, Clone, Copy, PartialEq, Eq)]
        struct Counted {
            touched: usize,
        }
        let commit = Commit::new((), (), Counted { touched: 2 });
        assert_eq!(commit.diagnostics().touched, 2);
    }

    #[test]
    fn commit_debug_does_not_format_owner_or_patch_payloads() {
        struct Payload;
        impl fmt::Debug for Payload {
            fn fmt(&self, _: &mut fmt::Formatter<'_>) -> fmt::Result {
                panic!("commit diagnostics must not format retained payloads")
            }
        }
        let commit = Commit::new(Payload, Payload, Diagnostics::CHANGED);
        assert_eq!(
            format!("{commit:?}"),
            "Commit { diagnostics: Diagnostics { changed: true }, .. }"
        );
    }

    #[test]
    fn refusal_kinds_have_unique_stable_names() {
        let names: std::collections::BTreeSet<_> =
            RefusalKind::ALL.iter().map(|kind| kind.as_str()).collect();
        assert_eq!(names.len(), RefusalKind::ALL.len());
        for kind in RefusalKind::ALL {
            assert_eq!(kind.to_string(), kind.as_str());
            assert!(
                kind.as_str()
                    .bytes()
                    .all(|byte| byte.is_ascii_lowercase() || byte == b'_')
            );
        }
    }

    #[test]
    fn refusal_displays_static_text_only() {
        let refusal = Refusal::new(
            RefusalKind::Protected,
            "set_text",
            "document or write protection is enforced",
        );
        assert_eq!(refusal.kind(), RefusalKind::Protected);
        assert_eq!(refusal.operation(), "set_text");
        assert_eq!(refusal.reason(), "document or write protection is enforced");
        assert_eq!(
            refusal.to_string(),
            "set_text refused (protected): document or write protection is enforced"
        );
        let error: &dyn std::error::Error = &refusal;
        assert!(error.source().is_none());
    }
}
