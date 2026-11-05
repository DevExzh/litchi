//! Source-backed RTF editing (record 0359a F2, the 0355a §1 shape).
//!
//! ```text
//! source_backed::Package::open(&rtf::Package, Arc<dyn ReadAt>, ExecutionContext)
//! Package::edit_tail_append(&self) -> tail_append::Edit
//! tail_append::Edit::append_*(&mut self, ..)            // staging only
//! tail_append::Edit::commit(self) -> Result<tail_append::Commit>   // validates, writes nothing
//! tail_append::Commit = litchi_core::edit::Commit<Package, tail_append::Patch, tail_append::Diagnostics>
//! Package::apply_tail_append(&self, &tail_append::Patch) -> Result<tail_append::Commit>
//! Package::write_to(&self, impl Write) / save(path)   // the only output
//! ```
//!
//! A [`Package`] is a positional RTF source bound, at open, to the parsed
//! document it was read as: the open proves the body tail is a plain
//! append point (the same conservative proof the eager tail append runs)
//! and binds the proof to the source's version, length and SHA-256 digest.
//! The package retains the proof and the source, never the parsed document
//! or a candidate artifact.
//!
//! A commit plans the bounded tail splice and returns a new [`Package`] that
//! holds the plan: [`Package::write_to`] and [`Package::save`] stream the
//! source around the inserted bytes through bounded windows, checking the
//! source's version and digest before the first and after the last write,
//! exactly as the former `TailAppendSourcePublicationPlan` did. An opened
//! package writes its source through the same checked identity stream.
//! Editing a committed package again materializes its output once (O(artifact)
//! memory for that package only) and proves the new tail against it.
//!
//! Every output is charged to the [`ExecutionContext`] the package was
//! opened with.

use std::fmt;
use std::io::{self, Write};
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_core::{ExecutionContext, OwnedSource, ReadAt, SourceVersion};

use crate::tail_append::{
    EditOptions, TailAppendPublicationError, TailAppendPublicationLimits, TailAppendSourceEdit,
    TailAppendSourceProof, TailAppendSourcePublicationPlan, TailSelector,
};

pub mod tail_append;

/// Options for opening a source-backed RTF package.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct OpenOptions {
    publication_limits: TailAppendPublicationLimits,
}

impl OpenOptions {
    /// The default bounded publication windows.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Select the bounded source and sink windows every output of the
    /// package (and of the packages its commits publish) uses.
    #[must_use]
    pub const fn publication_limits(mut self, limits: TailAppendPublicationLimits) -> Self {
        self.publication_limits = limits;
        self
    }

    /// The selected publication windows.
    #[must_use]
    pub const fn publication_limits_value(&self) -> TailAppendPublicationLimits {
        self.publication_limits
    }
}

/// An immutable source-backed RTF package; see the [module](self) docs.
#[derive(Clone)]
pub struct Package {
    inner: Arc<Inner>,
}

struct Inner {
    source: Arc<dyn ReadAt>,
    proof: TailAppendSourceProof,
    context: ExecutionContext,
    options: OpenOptions,
    /// The source parse profile to reuse when a committed output is
    /// materialized for a subsequent edit.
    document_options: crate::read::OpenOptions,
    /// The checked tail splice a committed package writes; `None` for an
    /// opened package, which writes its source.
    pending: Option<TailAppendSourcePublicationPlan>,
    /// A committed package's output, opened as a package of its own the
    /// first time the committed package is edited again.
    materialized: OnceLock<Package>,
}

impl fmt::Debug for Package {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Package")
            .field("source_bytes", &self.inner.proof.source_bytes())
            .field("source_version", &self.inner.proof.source_version())
            .field("committed", &self.inner.pending.is_some())
            .finish_non_exhaustive()
    }
}

impl Package {
    /// Open `source`, the exact bytes `document` was parsed from, with the
    /// default publication windows.
    ///
    /// # Errors
    ///
    /// Returns [`TailAppendPublicationError::Plan`] when the body tail is not
    /// a plain append point, and a source error when `source` does not hold
    /// exactly the document's bytes (version, length and digest), or the
    /// bounded digest scan exceeds `context`.
    pub fn open(
        document: &crate::Package,
        source: Arc<dyn ReadAt>,
        context: ExecutionContext,
    ) -> Result<Self, TailAppendPublicationError> {
        Self::open_with_options(document, source, context, OpenOptions::default())
    }

    /// [`Self::open`] with explicit options.
    ///
    /// # Errors
    ///
    /// As [`Self::open`], plus invalid publication windows.
    pub fn open_with_options(
        document: &crate::Package,
        source: Arc<dyn ReadAt>,
        context: ExecutionContext,
        options: OpenOptions,
    ) -> Result<Self, TailAppendPublicationError> {
        options.publication_limits.validate()?;
        let proof = TailAppendSourceProof::from_document_for_source_with_context(
            document,
            TailSelector::Body,
            source.as_ref(),
            &context,
        )?;
        Ok(Self::opened(
            source,
            proof,
            context,
            options,
            document.open_options(),
        ))
    }

    fn opened(
        source: Arc<dyn ReadAt>,
        proof: TailAppendSourceProof,
        context: ExecutionContext,
        options: OpenOptions,
        document_options: crate::read::OpenOptions,
    ) -> Self {
        Self {
            inner: Arc::new(Inner {
                source,
                proof,
                context,
                options,
                document_options,
                pending: None,
                materialized: OnceLock::new(),
            }),
        }
    }

    /// The package a family commit publishes: this package's source and
    /// proof with `plan` pending. Nothing is written.
    pub(crate) fn committed(&self, plan: TailAppendSourcePublicationPlan) -> Self {
        Self {
            inner: Arc::new(Inner {
                source: Arc::clone(&self.inner.source),
                proof: self.inner.proof.clone(),
                context: self.inner.context.clone(),
                options: self.inner.options,
                document_options: self.inner.document_options,
                pending: Some(plan),
                materialized: OnceLock::new(),
            }),
        }
    }

    /// Whether two handles are shares of one opened (or committed) package.
    pub(crate) fn same_package(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.inner, &other.inner)
    }

    /// Stream and authenticate the exact bytes this package would publish.
    /// The sink discards bytes, so this proof remains bounded by the
    /// publication plan's windows rather than materializing the artifact.
    pub(crate) fn replay_digest(&self) -> Result<(usize, String), TailAppendPublicationError> {
        let mut sink = io::sink();
        let report = self.write_to(&mut sink)?;
        Ok((report.bytes(), report.target_digest().to_owned()))
    }

    /// Rebase a validated source splice onto this package's source and
    /// publication policy.  The plan authenticates the recipient bytes before
    /// it is published, and the committed package retains this recipient's
    /// context and options.
    pub(crate) fn rebase_plan(
        &self,
        plan: &TailAppendSourcePublicationPlan,
        inverse: bool,
    ) -> Result<Self, TailAppendPublicationError> {
        let plan = plan.rebase(
            Arc::clone(&self.inner.source),
            self.inner.proof.clone(),
            self.inner.options.publication_limits,
            inverse,
        )?;
        Ok(self.committed(plan))
    }

    pub(crate) fn pending_plan(&self) -> Option<&TailAppendSourcePublicationPlan> {
        self.inner.pending.as_ref()
    }

    /// The options this package was opened with.
    #[must_use]
    pub fn options(&self) -> OpenOptions {
        self.inner.options
    }

    /// The source version the package's proof is bound to.
    #[must_use]
    pub fn source_version(&self) -> SourceVersion {
        self.inner.proof.source_version()
    }

    /// The exact source length the package's proof is bound to.
    #[must_use]
    pub fn source_bytes(&self) -> usize {
        self.inner.proof.source_bytes()
    }

    /// The exact length this package writes.
    #[must_use]
    pub fn len(&self) -> usize {
        self.inner.pending.as_ref().map_or_else(
            || self.inner.proof.source_bytes(),
            |plan| plan.output_bytes(),
        )
    }

    /// Whether this package writes no bytes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    /// The source-backed edit a family stages from: this package's source
    /// and proof, or (for a committed package) its materialized output's.
    pub(crate) fn staging(
        &self,
        options: EditOptions,
    ) -> Result<(Self, TailAppendSourceEdit), TailAppendPublicationError> {
        let base = self.edit_base()?;
        let edit = TailAppendSourceEdit::new_with_options(
            Arc::clone(&base.inner.source),
            base.inner.proof.clone(),
            TailSelector::Body,
            options,
        );
        Ok((base, edit))
    }

    fn edit_base(&self) -> Result<Self, TailAppendPublicationError> {
        if self.inner.pending.is_none() {
            return Ok(self.clone());
        }
        if let Some(materialized) = self.inner.materialized.get() {
            return Ok(materialized.clone());
        }
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(self.len()).map_err(|_error| {
            TailAppendPublicationError::Plan(crate::tail_append::Error::AllocationFailed {
                resource: "committed RTF output",
                requested: self.len(),
            })
        })?;
        self.write_to(&mut bytes)?;
        let document = crate::Package::from_bytes_with_options(&bytes, self.inner.document_options)
            .map_err(|error| {
                TailAppendPublicationError::Plan(crate::tail_append::Error::Rtf(error))
            })?;
        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
        let opened = Self::open_with_options(
            &document,
            source,
            self.inner.context.clone(),
            self.inner.options,
        )?;
        Ok(self.inner.materialized.get_or_init(|| opened).clone())
    }

    /// The publication context of this package.
    pub(crate) fn context(&self) -> &ExecutionContext {
        &self.inner.context
    }

    /// Streams this package's exact bytes: the planned tail splice for a
    /// committed package, the checked source for an opened one.
    ///
    /// # Errors
    ///
    /// Returns a source-changed or source-read error before the first write
    /// when the source no longer matches the proof, a sink error with its
    /// typed output progress, or a context refusal.
    pub fn write_to<W: Write>(&self, sink: &mut W) -> Result<Report, TailAppendPublicationError> {
        match &self.inner.pending {
            Some(plan) => plan.write_to(sink, &self.inner.context),
            None => self.identity_plan()?.write_to(sink, &self.inner.context),
        }
    }

    /// Atomically replaces `path` with this package's exact bytes, staged in
    /// a sibling temporary file with full durability.
    ///
    /// # Errors
    ///
    /// As [`Self::write_to`], plus the atomic-replacement errors.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<Report, OutputError> {
        self.save_with_options(path, crate::save::Options::new())
    }

    /// [`Self::save`] with explicit durability and overwrite policies.
    ///
    /// # Errors
    ///
    /// As [`Self::save`].
    pub fn save_with_options(
        &self,
        path: impl AsRef<Path>,
        options: crate::save::Options,
    ) -> Result<Report, OutputError> {
        litchi_core::atomic::save(
            path.as_ref(),
            options.durability_value(),
            options.overwrite_value(),
            |sink| self.write_to(sink).map_err(OutputError::Publication),
            || Ok(()),
        )
    }

    /// The checked identity stream of an opened package.
    fn identity_plan(&self) -> Result<TailAppendSourcePublicationPlan, TailAppendPublicationError> {
        TailAppendSourceEdit::new(
            Arc::clone(&self.inner.source),
            self.inner.proof.clone(),
            TailSelector::Body,
        )
        .publication_plan_with_context(self.inner.options.publication_limits, &self.inner.context)
    }
}

/// A source-backed RTF output failure.
#[derive(Debug)]
#[non_exhaustive]
pub enum OutputError {
    /// The checked publication failed; a sink failure keeps its typed
    /// output progress.
    Publication(TailAppendPublicationError),
    /// The atomic replacement of the destination failed.
    Save(litchi_core::atomic::Error),
}

impl fmt::Display for OutputError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Publication(error) => fmt::Display::fmt(error, formatter),
            Self::Save(error) => fmt::Display::fmt(error, formatter),
        }
    }
}

impl std::error::Error for OutputError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Publication(error) => Some(error),
            Self::Save(error) => Some(error),
        }
    }
}

impl From<litchi_core::atomic::Error> for OutputError {
    fn from(error: litchi_core::atomic::Error) -> Self {
        Self::Save(error)
    }
}

impl From<TailAppendPublicationError> for OutputError {
    fn from(error: TailAppendPublicationError) -> Self {
        Self::Publication(error)
    }
}

/// Evidence returned after source-backed publication succeeds.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Report {
    pub(crate) bytes: usize,
    pub(crate) source_bytes: usize,
    pub(crate) inserted_bytes: usize,
    pub(crate) source_version: SourceVersion,
    pub(crate) source_digest: String,
    pub(crate) target_digest: String,
    pub(crate) writes: usize,
    pub(crate) largest_write: usize,
}

impl Report {
    /// Complete target bytes accepted by the sink.
    #[must_use]
    pub const fn bytes(&self) -> usize {
        self.bytes
    }

    /// Exact source bytes emitted.
    #[must_use]
    pub const fn source_bytes(&self) -> usize {
        self.source_bytes
    }

    /// Exact inserted bytes emitted.
    #[must_use]
    pub const fn inserted_bytes(&self) -> usize {
        self.inserted_bytes
    }

    /// Source version checked before and after publication.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Canonical source digest checked before and after publication.
    #[must_use]
    pub fn source_digest(&self) -> &str {
        &self.source_digest
    }

    /// Canonical target digest computed during planning.
    #[must_use]
    pub fn target_digest(&self) -> &str {
        &self.target_digest
    }

    /// Number of sink writes.
    #[must_use]
    pub const fn writes(&self) -> usize {
        self.writes
    }

    /// Largest slice offered to one sink write.
    #[must_use]
    pub const fn largest_write(&self) -> usize {
        self.largest_write
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "unit tests panic on an unexpected result"
    )]

    use super::*;
    use std::num::{NonZeroU64, NonZeroUsize};

    fn context() -> ExecutionContext {
        let budget = litchi_core::Budget::root(
            "rtf-source-backed-unit",
            litchi_core::Limits::new(1 << 20, 1 << 20, 1 << 20, 1_000_000, 64, 1 << 20),
        );
        let (_source, token) = litchi_core::CancellationSource::pair();
        let limits = litchi_core::ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(64 * 1024).unwrap(),
            0,
        )
        .unwrap();
        ExecutionContext::new(budget, token, limits)
    }

    /// The family's committed output is byte-identical to the former
    /// `TailAppendSourceEdit` publication plan's for the same staging.
    #[test]
    fn family_output_matches_the_engine_plan() {
        let bytes = br#"{\rtf1\ansi A\par B}"#.to_vec();
        let document = crate::Package::from_bytes(&bytes).unwrap();
        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
        let context = context();

        let proof = TailAppendSourceProof::from_document_for_source(
            &document,
            TailSelector::Body,
            source.as_ref(),
        )
        .unwrap();
        let mut engine = TailAppendSourceEdit::new(Arc::clone(&source), proof, TailSelector::Body);
        engine.append_text_paragraphs(&["C", "D"]).unwrap();
        let plan = engine
            .publication_plan(TailAppendPublicationLimits::default())
            .unwrap();
        let mut expected = Vec::new();
        let expected_report = plan.write_to(&mut expected, &context).unwrap();

        let package = Package::open(&document, source, context).unwrap();
        let mut edit = package.edit_tail_append().unwrap();
        edit.append_text_paragraphs(&["C", "D"]).unwrap();
        let commit = edit.commit().unwrap();
        let mut output = Vec::new();
        let report = commit.package().write_to(&mut output).unwrap();
        assert_eq!(output, expected);
        assert_eq!(report.target_digest(), expected_report.target_digest());
        assert_eq!(report.bytes(), expected_report.bytes());
    }
}
