//! Semantic Pages section backgrounds and exact-source transactions.

use std::{fmt, sync::Arc};

use litchi_core::Position;
use litchi_iwa_common::color::Rgba;
use thiserror::Error as ThisError;

use super::Background;
use crate::Package;

/// A content-free semantic location used by transaction diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Pages package.
    Package,
    /// One section at its checked source position.
    Section { position: Position },
}

impl Path {
    /// Construct a path to one resolved section.
    #[must_use]
    pub const fn section(position: Position) -> Self {
        Self::Section { position }
    }

    /// Return the selected position, when this path names a section.
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

/// A finite resource governed by a section-background transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    InputBytes,
    OutputBytes,
    Entries,
    EntryBytes,
    TotalEntryBytes,
    PackageBytes,
    PayloadBytes,
    TotalPayloadBytes,
    PayloadObjects,
    PayloadMessages,
    PayloadItems,
    References,
    RetainedBytes,
    WireInputBytes,
    WireOutputBytes,
    WireFields,
    WireNesting,
    WireWork,
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

/// Failure from a semantic section-background read or transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    #[error(
        "the Pages section-background selector is ambiguous at positions {first:?} and {duplicate:?}"
    )]
    AmbiguousSelector {
        first: Position,
        duplicate: Position,
    },
    #[error("the Pages section-background selector did not match a section")]
    NameNotFound,
    #[error("the Pages section position {position:?} does not exist")]
    PositionNotFound { position: Position },
    #[error("the Pages {path} does not support exact section-background editing")]
    UnsupportedSource { path: Path },
    #[error("the Pages {path} has no unambiguous section background")]
    InvalidSource { path: Path },
    #[error(
        "Pages section-background {kind} limit exceeded at {path}: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        path: Path,
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Pages section-background transaction")]
    Allocation { amount: usize },
    #[error("the edited Pages {path} background failed verification")]
    Verification { path: Path },
    #[error("the Pages section-background patch does not match the exact source package")]
    PatchConflict,
}

/// One background staged against an immutable package snapshot.
#[must_use = "a Pages section-background edit must be committed"]
pub struct Edit<'a> {
    source: &'a Package,
    position: Position,
    before: Background,
    background: Background,
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
    /// Return the selected semantic location.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::section(self.position)
    }

    /// Borrow the background that would be published.
    #[must_use]
    pub const fn background(&self) -> &Background {
        &self.background
    }

    /// Stage a validated solid fill.
    pub fn set_solid(&mut self, color: Rgba) -> Result<&mut Self, Error> {
        self.background = Background::Solid(color);
        Ok(self)
    }

    /// Stage removal of the direct section fill.
    pub fn clear(&mut self) -> &mut Self {
        self.background = Background::None;
        self
    }

    /// Publish the staged value atomically.
    pub fn commit(self) -> Result<Commit, Error> {
        crate::package::section_background::commit_edit(self)
    }

    pub(crate) fn from_parts(source: &'a Package, position: Position, before: Background) -> Self {
        Self {
            source,
            position,
            background: before.clone(),
            before,
        }
    }

    pub(crate) fn into_parts(self) -> (&'a Package, Position, Background, Background) {
        (self.source, self.position, self.before, self.background)
    }
}

/// A reversible patch bound to exact source and target package artifacts.
#[derive(Clone, PartialEq)]
pub struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    position: Position,
    before: Background,
    after: Background,
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
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::section(self.position)
    }
    #[must_use]
    pub const fn before(&self) -> &Background {
        &self.before
    }
    #[must_use]
    pub const fn after(&self) -> &Background {
        &self.after
    }
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.source.as_ref() == self.target.as_ref()
    }
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
            touched_components: self.touched_components,
        }
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "crate-private exact patch handoff"
    )]
    pub(crate) fn from_parts(
        source: Arc<[u8]>,
        target: Arc<[u8]>,
        source_fingerprint: u64,
        target_fingerprint: u64,
        position: Position,
        before: Background,
        after: Background,
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
            touched_components,
        }
    }
    pub(crate) const fn position(&self) -> Position {
        self.position
    }
    pub(crate) const fn source_artifact(&self) -> &Arc<[u8]> {
        &self.source
    }
    pub(crate) const fn target_artifact(&self) -> &Arc<[u8]> {
        &self.target
    }
    pub(crate) const fn touched_components(&self) -> usize {
        self.touched_components
    }
}

/// Compact publication evidence without authored content or physical IDs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }
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
    pub(crate) const fn published(touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews: 0,
            full_reparse_performed: true,
        }
    }
}

/// The fully reopened result of one immutable background transaction.
#[must_use = "a Pages section-background commit contains a verified package snapshot"]
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
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
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
