//! Lossless Pages section settings and exact-source transactions.
//!
//! [`crate::section::Settings`] is the single semantic value for settings stored
//! directly on one Pages section. It preserves the native presence of the
//! optional name, four Boolean flags, and three pagination fields; absence is
//! never normalized to an explicit default. The transaction types in this
//! module retain no generated protobuf values and expose no native object
//! identifiers, package members, or raw source bytes.
//!
//! Select a section by its exact producer-visible name or checked source
//! position. Selection is resolved against the immutable source snapshot
//! before a replacement is staged:
//!
//! ```no_run
//! use litchi_pages::{Package, SectionSelector, section::Settings};
//!
//! # fn edit(source: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::from_bytes(source)?;
//! let selector = SectionSelector::name("Introduction");
//! let mut replacement: Settings = package.section_settings(selector)?;
//! replacement.set_first_page_hides_header_footer(Some(true));
//!
//! let commit = package
//!     .edit_section_settings(selector)?
//!     .set(replacement)?
//!     .commit()?;
//! let inverse = commit.patch().inverse();
//! let restored = commit.package().apply_section_settings(&inverse)?;
//! assert_eq!(restored.package().source_bytes(), package.source_bytes());
//! # Ok(())
//! # }
//! ```

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use thiserror::Error as ThisError;

use super::Settings;
use crate::Package;

/// A content-free semantic location used by transaction diagnostics.
///
/// Paths intentionally contain checked source positions rather than authored
/// section names or native object identifiers.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Pages package.
    Package,
    /// One section at its resolved zero-based semantic source position.
    Section {
        /// Checked position resolved against the immutable source snapshot.
        position: Position,
    },
}

impl Path {
    /// Construct a path to one resolved section.
    #[must_use]
    pub const fn section(position: Position) -> Self {
        Self::Section { position }
    }

    /// Return the selected section position, when this path names a section.
    #[must_use]
    pub const fn position(self) -> Option<Position> {
        match self {
            Self::Package => None,
            Self::Section { position } => Some(position),
        }
    }
}

impl fmt::Display for Path {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => formatter.write_str("package"),
            Self::Section { position } => write!(formatter, "section at {position:?}"),
        }
    }
}

/// A required native relationship missing from an otherwise valid source.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum DependencyKind {
    /// The previous section's template closure required for inheritance.
    PreviousSectionTemplates,
    /// The selected section's first-page template.
    FirstTemplate,
    /// The selected section's even-page template.
    EvenTemplate,
    /// The selected section's odd-page template.
    OddTemplate,
    /// The rooted derived layout-cache relationship.
    LayoutCache,
}

impl fmt::Display for DependencyKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PreviousSectionTemplates => "previous-section templates",
            Self::FirstTemplate => "first-page template",
            Self::EvenTemplate => "even-page template",
            Self::OddTemplate => "odd-page template",
            Self::LayoutCache => "layout cache",
        })
    }
}

/// A finite resource governed by a section-settings read or transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete edited package output bytes.
    OutputBytes,
    /// Package members, component objects, or component messages.
    Entries,
    /// Bytes retained by one package member.
    EntryBytes,
    /// Aggregate bytes retained by package members.
    TotalEntryBytes,
    /// Package names and structural metadata bytes.
    PackageBytes,
    /// Bytes in one decoded component payload.
    PayloadBytes,
    /// Aggregate decoded component payload bytes.
    TotalPayloadBytes,
    /// Component objects inspected by the transaction.
    PayloadObjects,
    /// Component messages inspected by the transaction.
    PayloadMessages,
    /// Component framing or metadata items inspected by the transaction.
    PayloadItems,
    /// Native references inspected while proving ownership.
    References,
    /// Aggregate semantic and transaction-owned retained bytes.
    RetainedBytes,
    /// Bytes inspected by the strict section-settings decoder.
    WireInputBytes,
    /// Bytes planned for one rewritten section payload.
    WireOutputBytes,
    /// Protobuf fields inspected or emitted.
    WireFields,
    /// Protobuf nesting inspected by the strict decoder.
    WireNesting,
    /// Aggregate strict decode and rewrite work.
    WireWork,
    /// Aggregate ownership, rewrite, reassembly, and reopen work.
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalEntryBytes => "total entry bytes",
            Self::PackageBytes => "package metadata bytes",
            Self::PayloadBytes => "component payload bytes",
            Self::TotalPayloadBytes => "total component payload bytes",
            Self::PayloadObjects => "component objects",
            Self::PayloadMessages => "component messages",
            Self::PayloadItems => "component metadata items",
            Self::References => "references",
            Self::RetainedBytes => "retained bytes",
            Self::WireInputBytes => "wire input bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::TransactionWork => "transaction work",
        })
    }
}

/// Failure from a semantic section-settings read or immutable transaction.
///
/// `Display` and `Debug` output omit authored section names, field values,
/// native identifiers, package member names, raw bytes, and retained patch
/// artifacts. [`enum@crate::section::Error`] is the shared section semantic validator; an
/// aggregate settings value can produce only its name and pagination variants.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// More than one section matched an exact-name selector.
    #[error(
        "the Pages section-settings selector is ambiguous at positions {first:?} and {duplicate:?}"
    )]
    AmbiguousSelector {
        /// First matching semantic source position.
        first: Position,
        /// Next matching semantic source position.
        duplicate: Position,
    },
    /// No section matched an exact-name selector.
    #[error("the Pages section-settings selector did not match a section")]
    NameNotFound,
    /// No section exists at the requested semantic source position.
    #[error("the Pages section position {position:?} does not exist")]
    PositionNotFound {
        /// Requested checked source position.
        position: Position,
    },
    /// The replacement violates an archive-free section invariant.
    #[error("invalid Pages section settings: {0}")]
    InvalidSettings(#[source] super::Error),
    /// The physical source cannot publish a preservation-safe changed edit.
    #[error("the Pages {path} does not support exact section-settings editing")]
    UnsupportedSource {
        /// Content-free semantic location of the unsupported owner.
        path: Path,
    },
    /// A required, otherwise valid native relationship is absent.
    #[error("the Pages {path} is missing its required {kind}")]
    UnsupportedDependency {
        /// Content-free semantic location requiring the relationship.
        path: Path,
        /// Kind of missing relationship.
        kind: DependencyKind,
    },
    /// The rooted native owner is malformed or ambiguous.
    #[error("the Pages {path} has no unambiguous editable section settings")]
    InvalidSource {
        /// Content-free semantic location of the invalid owner.
        path: Path,
    },
    /// A finite transaction resource ceiling was exceeded.
    #[error(
        "Pages section-settings {kind} limit exceeded at {path}: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Content-free semantic location where the ceiling was exceeded.
        path: Path,
        /// Resource category that exceeded its ceiling.
        kind: LimitKind,
        /// Observed or requested resource amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded transaction allocation failed before publication.
    #[error("could not allocate {amount} units for the Pages section-settings transaction")]
    Allocation {
        /// Requested bytes or elements.
        amount: usize,
    },
    /// Complete candidate reopening or locality checking failed.
    #[error("the edited Pages {path} settings failed verification")]
    Verification {
        /// Content-free semantic location that failed verification.
        path: Path,
    },
    /// The supplied patch was not created from this exact package artifact.
    #[error("the Pages section-settings patch does not match the exact source package")]
    PatchConflict,
}

/// One complete section-settings value staged against an immutable package.
#[must_use = "a Pages section-settings edit must be committed to publish its staged value"]
pub struct Edit<'a> {
    source: &'a Package,
    position: Position,
    before: Settings,
    settings: Settings,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.path())
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    /// Return the resolved, content-free semantic target.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::section(self.position)
    }

    /// Borrow the complete lossless value that would be published.
    #[must_use]
    pub const fn settings(&self) -> &Settings {
        &self.settings
    }

    /// Replace the complete staged value, returning this edit for chaining.
    ///
    /// The replacement is validated before it is retained, so a failure
    /// leaves the immutable source package untouched.
    pub fn set(self, settings: Settings) -> Result<Self, Error> {
        settings.validate().map_err(Error::InvalidSettings)?;
        Ok(self.replace_settings_unchecked(settings))
    }

    /// Atomically publish the staged value after exact-source verification.
    pub fn commit(self) -> Result<Commit, Error> {
        crate::package::section_settings::commit_edit(self)
    }

    pub(crate) fn from_package_parts(
        source: &'a Package,
        position: Position,
        before: Settings,
    ) -> Self {
        Self {
            source,
            position,
            settings: before.clone(),
            before,
        }
    }

    pub(crate) fn into_package_parts(self) -> (&'a Package, Position, Settings, Settings) {
        (self.source, self.position, self.before, self.settings)
    }

    pub(crate) fn replace_settings_unchecked(mut self, settings: Settings) -> Self {
        self.settings = settings;
        self
    }
}

/// A reversible patch bound to exact source and target package artifacts.
///
/// Exact bytes and native locality proofs remain private. Compact fingerprints
/// are diagnostics only and never authorize application.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    position: Position,
    before: Settings,
    after: Settings,
    source_layout_state: Option<u64>,
    target_layout_state: Option<u64>,
    source_preview_count: usize,
    target_preview_count: usize,
    touched_components: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("path", &self.path())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the resolved, content-free semantic target.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::section(self.position)
    }

    /// Borrow the complete settings required from the patch source.
    #[must_use]
    pub const fn before(&self) -> &Settings {
        &self.before
    }

    /// Borrow the complete settings produced by the patch target.
    #[must_use]
    pub const fn after(&self) -> &Settings {
        &self.after
    }

    /// Return the source artifact's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the target artifact's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return whether semantic state and exact package bytes are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && self.source.as_ref() == self.target.as_ref()
    }

    /// Return the exact reverse direction from target back to source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            position: self.position,
            before: self.after.clone(),
            after: self.before.clone(),
            source_layout_state: self.target_layout_state,
            target_layout_state: self.source_layout_state,
            source_preview_count: self.target_preview_count,
            target_preview_count: self.source_preview_count,
            touched_components: self.touched_components,
        }
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "crate-private exact patch handoff"
    )]
    pub(crate) fn from_package_parts(
        source: Arc<[u8]>,
        target: Arc<[u8]>,
        source_fingerprint: u64,
        target_fingerprint: u64,
        position: Position,
        before: Settings,
        after: Settings,
        source_layout_state: Option<u64>,
        target_layout_state: Option<u64>,
        source_preview_count: usize,
        target_preview_count: usize,
        touched_components: usize,
    ) -> Self {
        Self {
            source,
            target,
            source_fingerprint,
            target_fingerprint,
            position,
            before,
            after,
            source_layout_state,
            target_layout_state,
            source_preview_count,
            target_preview_count,
            touched_components,
        }
    }

    pub(crate) fn source_artifact(&self) -> &Arc<[u8]> {
        &self.source
    }

    pub(crate) fn target_artifact(&self) -> &Arc<[u8]> {
        &self.target
    }

    pub(crate) const fn position(&self) -> Position {
        self.position
    }

    pub(crate) const fn source_layout_state(&self) -> Option<u64> {
        self.source_layout_state
    }

    pub(crate) const fn target_layout_state(&self) -> Option<u64> {
        self.target_layout_state
    }

    pub(crate) const fn source_preview_count(&self) -> usize {
        self.source_preview_count
    }

    pub(crate) const fn target_preview_count(&self) -> usize {
        self.target_preview_count
    }

    pub(crate) const fn touched_components(&self) -> usize {
        self.touched_components
    }
}

/// Compact, content-free evidence for one section-settings commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    /// Return whether the committed artifact differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of physical IWA components rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return the number of canonical root previews removed.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }

    pub(crate) const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    pub(crate) const fn published(
        touched_components: usize,
        deleted_previews: usize,
        full_reparse_performed: bool,
    ) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed,
        }
    }
}

/// The fully verified result of one immutable section-settings transaction.
#[must_use = "a Pages section-settings commit contains the verified package snapshot"]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("path", &self.patch.path())
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}

impl Commit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    pub(crate) const fn from_parts(
        package: Package,
        patch: Patch,
        diagnostics: Diagnostics,
    ) -> Self {
        Self {
            package,
            patch,
            diagnostics,
        }
    }
}
