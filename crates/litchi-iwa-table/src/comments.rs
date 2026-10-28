//! Format-neutral transaction kernel for table-cell comments and direct replies.
//!
//! The three iWork leaves store table comments in the same TST sidecar graph,
//! but each owner resolves its semantic table selector and reopens its own
//! package type.  This module owns the archive-free semantic state, operation
//! algebra, exact-source patch contract, and the bounded native graph
//! admission walk.  A leaf supplies the format-specific selector, wire
//! rewrite and candidate publication, then maps its private failures into
//! [`Error`].
//!
//! Native object identifiers, list keys, storage UUIDs, protobuf values, and
//! archive member names are intentionally confined to the private leaf
//! adapters.  The module is public only to allow the three leaf crates to
//! consume the shared implementation; it is not re-exported by a supported
//! facade.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The comment transaction keeps its algebra beside its bounded validation helpers."
)]

use std::collections::{HashSet, VecDeque};
use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_common::table::read::{Comment, CommentReply};
use thiserror::Error as ThisError;

use crate::structure::{Fault, TableLeaf};

type CellPosition = At;

/// The largest semantic comment or reply body accepted by the shared kernel.
pub const MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;

/// The bounded native graph census performed before a leaf rewrites a
/// comment thread.
///
/// The public facades resolve their format-specific table selectors and own
/// ZIP publication, but they must not each invent a different admission
/// walk.  This result is deliberately small: it is the common proof that all
/// rooted table objects are present, that references can be traversed within
/// the package ceilings, and that the rewrite has a finite amount of graph
/// work to charge before it starts.  A missing target of a non-root edge is
/// recorded as an external reference; opaque iWork graphs legitimately carry
/// such edges, and the leaf writer remains responsible for deciding whether a
/// particular edge is in its mutation profile.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct GraphAdmission {
    objects: usize,
    messages: usize,
    references: usize,
    external_references: usize,
    maximum_depth: usize,
    work: usize,
    allocations: usize,
}

impl GraphAdmission {
    /// Number of native objects reached from the selected roots.
    #[must_use]
    pub const fn objects(self) -> usize {
        self.objects
    }

    /// Number of native messages in the reached objects.
    #[must_use]
    pub const fn messages(self) -> usize {
        self.messages
    }

    /// Number of object/data references inspected.
    #[must_use]
    pub const fn references(self) -> usize {
        self.references
    }

    /// Number of references whose target is outside the decoded catalog.
    #[must_use]
    pub const fn external_references(self) -> usize {
        self.external_references
    }

    /// Maximum edge depth reached by the bounded walk.
    #[must_use]
    pub const fn maximum_depth(self) -> usize {
        self.maximum_depth
    }

    /// Conservative work units to charge to the leaf's transaction ledger.
    #[must_use]
    pub const fn work(self) -> usize {
        self.work
    }

    /// Allocation units required by the bounded queue and visited set.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
}

/// Walk the native objects rooted by one table selection before mutation.
///
/// This is the shared graph algorithm for table-cell comment routes.  It is
/// intentionally independent of a Pages/Keynote/Numbers selector and does
/// not expose native identifiers through a supported facade.  Roots must be
/// non-zero object identifiers already proved by the leaf selector resolver.
/// The walk follows both object and data references, de-duplicates object
/// identities, detects missing roots, and admits every queue/reference step
/// against the source's semantic ceilings before returning.
pub fn admit_graph(
    source: &impl crate::TableSource,
    roots: &[u64],
) -> Result<GraphAdmission, Error> {
    const MAX_DEPTH: usize = 128;
    let max_objects = source.max_objects().max(1);
    let max_references = source.max_references().max(1);
    let mut queue = VecDeque::new();
    let mut seen = HashSet::new();
    queue
        .try_reserve(roots.len())
        .map_err(|_| Error::Allocation {
            amount: roots.len(),
        })?;
    seen.try_reserve(roots.len())
        .map_err(|_| Error::Allocation {
            amount: roots.len(),
        })?;
    let initial_allocations = roots.len().checked_mul(2).ok_or(Error::InvalidSource)?;
    for &root in roots {
        if root == 0 {
            return Err(Error::InvalidSource);
        }
        if source
            .resolve_object(root)
            .map_err(|_| Error::InvalidSource)?
            .is_none()
        {
            return Err(Error::InvalidSource);
        }
        if seen.insert(root) {
            queue.push_back((root, 0usize));
        }
    }

    let mut admission = GraphAdmission {
        allocations: initial_allocations,
        ..GraphAdmission::default()
    };
    while let Some((identifier, depth)) = queue.pop_front() {
        let resolved = source
            .resolve_object(identifier)
            .map_err(|_| Error::InvalidSource)?
            .ok_or(Error::InvalidSource)?;
        let component = source
            .catalog()
            .get_index(resolved.component_index)
            .ok_or(Error::InvalidSource)?;
        let object = component
            .archive()
            .objects
            .get(resolved.object_index)
            .ok_or(Error::InvalidSource)?;
        if object.archive_info.identifier != Some(identifier) {
            return Err(Error::InvalidSource);
        }
        admission.objects = admission
            .objects
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;
        if admission.objects > max_objects {
            return Err(Error::Limit {
                kind: LimitKind::WireWork,
                observed: admission.objects as u64,
                maximum: max_objects as u64,
            });
        }
        admission.maximum_depth = admission.maximum_depth.max(depth);
        if depth >= MAX_DEPTH {
            return Err(Error::Limit {
                kind: LimitKind::WireWork,
                observed: (depth + 1) as u64,
                maximum: MAX_DEPTH as u64,
            });
        }

        admission.messages = admission
            .messages
            .checked_add(object.messages.len())
            .ok_or(Error::InvalidSource)?;
        admission.work = admission
            .work
            .checked_add(
                size_of::<litchi_iwa_core::ArchiveObject>()
                    .checked_add(
                        object
                            .messages
                            .len()
                            .checked_mul(size_of::<litchi_iwa_core::RawMessage>())
                            .ok_or(Error::InvalidSource)?,
                    )
                    .ok_or(Error::InvalidSource)?,
            )
            .ok_or(Error::InvalidSource)?;
        admission.allocations = admission
            .allocations
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;

        let mut visit_reference = |target: u64| -> Result<(), Error> {
            if target == 0 {
                return Err(Error::InvalidSource);
            }
            admission.references = admission
                .references
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
            if admission.references > max_references {
                return Err(Error::Limit {
                    kind: LimitKind::References,
                    observed: admission.references as u64,
                    maximum: max_references as u64,
                });
            }
            admission.work = admission
                .work
                .checked_add(size_of::<u64>())
                .ok_or(Error::InvalidSource)?;
            let Some(_) = source
                .resolve_object(target)
                .map_err(|_| Error::InvalidSource)?
            else {
                admission.external_references = admission
                    .external_references
                    .checked_add(1)
                    .ok_or(Error::InvalidSource)?;
                return Ok(());
            };
            if seen.contains(&target) {
                return Ok(());
            }
            seen.try_reserve(1)
                .map_err(|_| Error::Allocation { amount: 1 })?;
            queue
                .try_reserve(1)
                .map_err(|_| Error::Allocation { amount: 1 })?;
            admission.allocations = admission
                .allocations
                .checked_add(2)
                .ok_or(Error::InvalidSource)?;
            seen.insert(target);
            queue.push_back((target, depth + 1));
            Ok(())
        };

        for info in &object.archive_info.message_infos {
            let field_work = info
                .field_infos
                .iter()
                .try_fold(
                    size_of::<litchi_iwa_core::archive::MessageInfo>(),
                    |work, field| {
                        work.checked_add(size_of::<litchi_iwa_core::archive::FieldInfo>())
                            .and_then(|work| {
                                work.checked_add(
                                    field.path.path.len().checked_mul(size_of::<u32>())?,
                                )
                            })
                    },
                )
                .ok_or(Error::InvalidSource)?;
            admission.work = admission
                .work
                .checked_add(field_work)
                .ok_or(Error::InvalidSource)?;
            for &reference in &info.object_references {
                visit_reference(reference)?;
            }
            for &reference in &info.data_references {
                visit_reference(reference)?;
            }
            for field in &info.field_infos {
                for &reference in &field.object_references {
                    visit_reference(reference)?;
                }
                for &reference in &field.data_references {
                    visit_reference(reference)?;
                }
            }
        }
    }
    Ok(admission)
}

/// A content-free semantic location associated with a comment operation.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete package.
    Package,
    /// One table cell.  The table and cell are semantic positions only.
    Cell {
        scope: usize,
        table: usize,
        cell: CellPosition,
    },
}

impl fmt::Debug for Path {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => formatter.write_str("Path::Package"),
            Self::Cell { .. } => formatter.write_str("Path::Cell"),
        }
    }
}

/// Finite resources charged by one comment transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source bytes retained by the transaction.
    InputBytes,
    /// Complete candidate bytes retained by the transaction.
    OutputBytes,
    /// Native payload bytes inspected or emitted.
    PayloadBytes,
    /// Native references inspected.
    References,
    /// Reply records inspected.
    Replies,
    /// Comment and reply text bytes retained.
    TextBytes,
    /// Wire fields inspected.
    WireFields,
    /// Wire work performed.
    WireWork,
    /// Bounded scratch or allocation work.
    Allocations,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::PayloadBytes => "payload bytes",
            Self::References => "references",
            Self::Replies => "replies",
            Self::TextBytes => "text bytes",
            Self::WireFields => "wire fields",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
        })
    }
}

/// A well-formed request that the focused comment route refuses.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The semantic selector did not identify a table.
    #[error("the selected table was not found")]
    TableNotFound,
    /// The semantic selector identified more than one table.
    #[error("the selected table is ambiguous")]
    AmbiguousSelector,
    /// The source is not an exact package owned by this transaction.
    #[error("the source does not support exact table-comment editing")]
    UnsupportedSource,
    /// The selected table is protected by its native lock state.
    #[error("the selected table is locked")]
    TableLocked,
    /// The comment graph has an edge outside the focused writer profile.
    #[error("the table-comment graph has an unsupported dependency")]
    UnsupportedDependency,
    /// A root comment is required by the requested reply operation.
    #[error("the selected cell has no comment")]
    CommentNotFound,
    /// The requested reply ordinal is absent.
    #[error("the selected comment reply does not exist")]
    ReplyNotFound { index: usize },
}

impl Refusal {
    /// Classify the refusal for the unified operations layer.
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound | Self::CommentNotFound | Self::ReplyNotFound { .. } => {
                RefusalKind::NotFound
            },
            Self::AmbiguousSelector => RefusalKind::InvalidValue,
            Self::UnsupportedSource | Self::UnsupportedDependency => {
                RefusalKind::UnsupportedContent
            },
            Self::TableLocked => RefusalKind::Protected,
        }
    }
}

/// A terminal failure from the shared comment transaction.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request is outside this focused route.
    #[error(transparent)]
    Refused(Refusal),
    /// The selected native graph is malformed.
    #[error("the table-comment source is invalid")]
    InvalidSource,
    /// A finite transaction resource was exceeded.
    #[error("table comments {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded staging allocation failed.
    #[error("could not allocate {amount} units for table comments")]
    Allocation { amount: usize },
    /// The reopened candidate did not reproduce the requested state.
    #[error("the edited table comment failed semantic verification")]
    Verification,
    /// A patch was applied to a changed or different exact source.
    #[error("the table-comment patch does not match the exact source")]
    PatchConflict,
}

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Unsupported => Self::Refused(Refusal::UnsupportedSource),
            Fault::Invalid | Fault::Dependency => Self::InvalidSource,
            Fault::Limit {
                observed, maximum, ..
            } => Self::Limit {
                kind: LimitKind::WireWork,
                observed,
                maximum,
            },
            Fault::Allocation { amount } => Self::Allocation { amount },
            Fault::Verification => Self::Verification,
        }
    }
}

/// The complete semantic state of one selected cell's comment thread.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    comment: Option<Comment>,
    replies: Box<[CommentReply]>,
}

impl Snapshot {
    /// Construct a root/reply snapshot after the leaf has charged its read.
    #[must_use]
    pub fn new(comment: Option<Comment>, replies: impl Into<Box<[CommentReply]>>) -> Self {
        Self {
            comment,
            replies: replies.into(),
        }
    }

    /// Return the root comment, if one is attached.
    #[must_use]
    pub fn comment(&self) -> Option<&Comment> {
        self.comment.as_ref()
    }

    /// Return the direct replies resolved by the leaf.
    #[must_use]
    pub fn replies(&self) -> &[CommentReply] {
        &self.replies
    }

    /// Consume the snapshot into its semantic parts.
    #[must_use]
    pub fn into_parts(self) -> (Option<Comment>, Box<[CommentReply]>) {
        (self.comment, self.replies)
    }
}

/// One bounded semantic operation over a comment thread.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Operation {
    /// Replace the root body, preserving source metadata where supported.
    Set(Box<str>),
    /// Remove the root and its direct reply thread.
    Clear,
    /// Append one direct reply.
    AddReply(Box<str>),
    /// Replace one direct reply by ordinal.
    SetReply { index: usize, text: Box<str> },
    /// Remove one direct reply by ordinal.
    RemoveReply { index: usize },
}

/// Apply the semantic operation before any native rewrite is attempted.
pub fn apply_operation(before: &Snapshot, operation: &Operation) -> Result<Snapshot, Error> {
    let mut comment = before.comment.clone();
    let mut replies = before.replies.to_vec();
    let projects_replies = comment
        .as_ref()
        .is_some_and(|comment| comment.replies().is_some());
    match operation {
        Operation::Set(text) => {
            comment = Some(match comment.as_ref() {
                Some(existing) if existing.replies().is_some() => Comment::with_replies(
                    text.as_ref(),
                    existing.timestamp(),
                    existing.author().cloned(),
                    existing.replies().into_iter().flatten().cloned(),
                ),
                Some(existing) => Comment::with_metadata(
                    text.as_ref(),
                    existing.timestamp(),
                    existing.author().cloned(),
                ),
                None => Comment::new(text.as_ref()),
            });
        },
        Operation::Clear => {
            if comment.is_none() {
                return Err(Error::Refused(Refusal::CommentNotFound));
            }
            comment = None;
            replies.clear();
        },
        Operation::AddReply(text) => {
            if comment.is_none() {
                return Err(Error::Refused(Refusal::CommentNotFound));
            }
            replies
                .try_reserve(1)
                .map_err(|_| Error::Allocation { amount: 1 })?;
            replies.push(CommentReply::new(text.as_ref()));
        },
        Operation::SetReply { index, text } => {
            let Some(reply) = replies.get_mut(*index) else {
                return Err(Error::Refused(Refusal::ReplyNotFound { index: *index }));
            };
            let (timestamp, author) = (reply.timestamp(), reply.author().cloned());
            *reply = CommentReply::with_metadata(text.as_ref(), timestamp, author);
        },
        Operation::RemoveReply { index } => {
            if *index >= replies.len() {
                return Err(Error::Refused(Refusal::ReplyNotFound { index: *index }));
            }
            replies.remove(*index);
        },
    }
    if projects_replies {
        if let Some(existing) = comment.take() {
            comment = Some(Comment::with_replies(
                existing.text().to_owned(),
                existing.timestamp(),
                existing.author().cloned(),
                replies.iter().cloned(),
            ));
        }
    }
    Ok(Snapshot::new(comment, replies.into_boxed_slice()))
}

/// An exact-source, reversible semantic comment patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    path: Path,
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Create a patch after the leaf has retained both exact artifacts.
    pub(crate) fn new(
        artifacts: ExactArtifacts,
        path: Path,
        before: Snapshot,
        after: Snapshot,
    ) -> Self {
        Self {
            artifacts,
            path,
            before,
            after,
        }
    }

    /// Return the semantic location of the patch.
    #[must_use]
    pub const fn path(&self) -> Path {
        self.path
    }

    /// Borrow the source semantic state.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the target semantic state.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is a semantic and byte-identical no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after || self.artifacts.source_bytes() == self.artifacts.target_bytes()
    }

    /// Return the exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            path: self.path,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Test whether this patch authorizes the package's current exact source.
    #[must_use]
    pub fn authorizes(&self, source: &impl crate::TableSource) -> bool {
        self.artifacts.source().as_ref() == source.source_bytes()
    }

    /// Return the retained exact target bytes for a leaf reopener.
    #[must_use]
    pub(crate) fn target_bytes(&self) -> Arc<[u8]> {
        self.artifacts.target()
    }
}

/// Publication diagnostics shared by the three leaves.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Diagnostics {
    touched_components: usize,
    changed: bool,
}

impl Diagnostics {
    /// Construct diagnostics for an unchanged source.
    #[must_use]
    pub const fn unchanged() -> Self {
        Self {
            touched_components: 0,
            changed: false,
        }
    }

    /// Construct diagnostics for a changed publication.
    #[must_use]
    pub const fn changed(touched_components: usize) -> Self {
        Self {
            touched_components,
            changed: true,
        }
    }

    /// Return the number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the published package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// A validated package snapshot and its exact-source patch.
#[must_use = "a comment commit contains the verified package snapshot"]
pub struct Commit<P> {
    package: P,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl<P> Commit<P> {
    /// Construct a commit from the leaf's reopened candidate.
    pub(crate) fn new(package: P, patch: Patch, diagnostics: Diagnostics) -> Self {
        Self {
            package,
            patch,
            diagnostics,
        }
    }

    /// Borrow the reopened package.
    #[must_use]
    pub const fn package(&self) -> &P {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> P {
        self.package
    }

    /// Consume the verified commit and return its package and exact patch.
    #[must_use]
    pub fn into_parts(self) -> (P, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }
}

/// Validate a candidate through the leaf's strict ingress and semantic read.
///
/// The native adapter must charge its graph rewrite before calling this
/// helper.  The helper then retains the candidate artifact, reopens it once,
/// and compares the requested semantic state without permitting a host
/// fallback.
pub fn finish<P, F>(
    source: &P,
    path: Path,
    before: Snapshot,
    after: Snapshot,
    candidate: Arc<[u8]>,
    touched_components: usize,
    readback: F,
) -> Result<Commit<P>, Error>
where
    P: TableLeaf + Clone,
    F: FnOnce(&P) -> Result<Snapshot, Error>,
{
    let source_bytes: Arc<[u8]> = Arc::from(source.source_bytes());
    if source_bytes.as_ref() == candidate.as_ref() {
        if before != after {
            return Err(Error::Verification);
        }
        let artifacts = ExactArtifacts::new(source_bytes, candidate);
        return Ok(Commit::new(
            source.clone(),
            Patch::new(artifacts, path, before, after),
            Diagnostics::unchanged(),
        ));
    }

    let reopened = source
        .reopen_candidate(Arc::clone(&candidate))
        .map_err(Error::from)?;
    let observed = readback(&reopened)?;
    if observed != after {
        return Err(Error::Verification);
    }
    let artifacts = ExactArtifacts::new(source_bytes, candidate);
    Ok(Commit::new(
        reopened,
        Patch::new(artifacts, path, before, after),
        Diagnostics::changed(touched_components),
    ))
}

/// A bounded text copy used by leaf edits before native work begins.
pub fn copy_text(text: &str, maximum: usize) -> Result<Box<str>, Error> {
    if text.len() > maximum.min(MAX_TEXT_BYTES) {
        return Err(Error::Limit {
            kind: LimitKind::TextBytes,
            observed: text.len() as u64,
            maximum: maximum.min(MAX_TEXT_BYTES) as u64,
        });
    }
    let mut owned = String::new();
    owned
        .try_reserve_exact(text.len())
        .map_err(|_| Error::Allocation { amount: text.len() })?;
    owned.push_str(text);
    Ok(owned.into_boxed_str())
}

#[cfg(test)]
mod tests {
    use super::*;

    struct EmptySource {
        catalog: litchi_iwa_archive::ComponentCatalog,
    }

    impl crate::TableSource for EmptySource {
        fn catalog(&self) -> &litchi_iwa_archive::ComponentCatalog {
            &self.catalog
        }

        fn physical(&self) -> Option<&litchi_iwa_archive::SourceCatalog> {
            None
        }

        fn resolve_object(
            &self,
            _identifier: u64,
        ) -> Result<Option<crate::ResolvedObject<'_>>, crate::ResolveError> {
            Ok(None)
        }

        fn source_bytes(&self) -> &[u8] {
            &[]
        }

        fn archive_limits(&self) -> litchi_iwa_archive::Limits {
            litchi_iwa_archive::Limits::default()
        }

        fn max_objects(&self) -> usize {
            1
        }

        fn max_references(&self) -> usize {
            1
        }
    }

    #[test]
    fn graph_admission_requires_resolved_roots() {
        let source = EmptySource {
            catalog: litchi_iwa_archive::ComponentCatalog::__empty(),
        };
        assert_eq!(admit_graph(&source, &[]), Ok(GraphAdmission::default()));
        assert_eq!(admit_graph(&source, &[1]), Err(Error::InvalidSource));
    }

    #[test]
    fn operation_preserves_metadata_and_rejects_missing_reply() {
        let before = Snapshot::new(
            Some(Comment::with_metadata("root", None, None)),
            vec![CommentReply::new("reply")].into_boxed_slice(),
        );
        let after = apply_operation(
            &before,
            &Operation::SetReply {
                index: 0,
                text: "updated".into(),
            },
        )
        .unwrap();
        assert_eq!(after.comment().unwrap().text(), "root");
        assert_eq!(after.replies()[0].text(), "updated");
        assert!(matches!(
            apply_operation(&before, &Operation::RemoveReply { index: 2 },),
            Err(Error::Refused(Refusal::ReplyNotFound { index: 2 }))
        ));
    }

    #[test]
    fn clear_removes_replies_and_empty_root_is_a_noop_only_when_unchanged() {
        let before = Snapshot::new(
            Some(Comment::new("root")),
            vec![CommentReply::new("reply")].into_boxed_slice(),
        );
        let after = apply_operation(&before, &Operation::Clear).unwrap();
        assert!(after.comment().is_none());
        assert!(after.replies().is_empty());
        assert!(matches!(
            apply_operation(&after, &Operation::Clear),
            Err(Error::Refused(Refusal::CommentNotFound))
        ));
    }
}
