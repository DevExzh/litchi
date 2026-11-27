//! The source-backed tail-append family: plain paragraphs appended at the
//! proven body tail of a source-backed RTF package, published as a bounded
//! tail splice (the former `TailAppendSourceEdit` publisher).

use std::fmt;

use super::Package;
use crate::tail_append::{
    EditOptions, PlainParagraph, PlainRun, TailAppendSourceEdit, TailAppendSourcePublicationPlan,
};

/// The family error: the checked source-backed publication error.
pub use crate::tail_append::TailAppendPublicationError as Error;

/// The finite append dimensions an [`Edit`] is bounded by (paragraphs, runs,
/// text bytes; see [`EditOptions`]).
pub use crate::tail_append::LimitKind;

/// A published source-backed tail append: the committed package, its
/// exact-pair patch and diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch, Diagnostics>;

/// A staged source-backed tail append. Dropping it discards it.
pub struct Edit {
    base: Package,
    inner: TailAppendSourceEdit,
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("paragraphs", &self.inner.paragraph_count())
            .field("runs", &self.inner.run_count())
            .finish_non_exhaustive()
    }
}

impl Edit {
    /// Stage one bounded batch of borrowed plain paragraphs.
    ///
    /// # Errors
    ///
    /// Returns a bound or text refusal, or an unsupported-operation refusal for a second
    /// batch; the staged state is unchanged.
    pub fn append_paragraphs(
        &mut self,
        paragraphs: &[PlainParagraph<'_>],
    ) -> Result<&mut Self, Error> {
        self.inner
            .append_paragraphs(paragraphs)
            .map_err(Error::Plan)?;
        Ok(self)
    }

    /// Stage one plain run per paragraph.
    ///
    /// # Errors
    ///
    /// As [`Self::append_paragraphs`].
    pub fn append_text_paragraphs(&mut self, paragraphs: &[&str]) -> Result<&mut Self, Error> {
        self.inner
            .append_text_paragraphs(paragraphs)
            .map_err(Error::Plan)?;
        Ok(self)
    }

    /// Stage one paragraph made from plain runs.
    ///
    /// # Errors
    ///
    /// As [`Self::append_paragraphs`].
    pub fn append_runs(&mut self, runs: &[PlainRun<'_>]) -> Result<&mut Self, Error> {
        self.inner.append_runs(runs).map_err(Error::Plan)?;
        Ok(self)
    }

    /// The append bounds this edit was started with.
    #[must_use]
    pub fn options(&self) -> EditOptions {
        self.inner.options()
    }

    /// Validate the staged paragraphs against the package's source and plan
    /// the bounded tail splice; nothing is written. The committed package
    /// writes the plan.
    ///
    /// # Errors
    ///
    /// Returns a source-changed or source-read error, a plan refusal, or a
    /// context refusal of the bounded planning scan.
    pub fn commit(self) -> Result<Commit, Error> {
        let plan = self.inner.publication_plan_with_context(
            self.base.options().publication_limits_value(),
            self.base.context(),
        )?;
        let diagnostics = Diagnostics::of(&plan);
        let after = if plan.is_noop() {
            self.base.clone()
        } else {
            self.base.committed(plan)
        };
        let patch = Patch {
            before: self.base,
            after: after.clone(),
            diagnostics: diagnostics.clone(),
        };
        Ok(Commit::new(after, patch, diagnostics))
    }
}

/// Content-free facts about one source-backed tail-append commit.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    source_bytes: usize,
    inserted_bytes: usize,
    output_bytes: usize,
    input_bytes: usize,
    paragraphs: usize,
    runs: usize,
    source_digest: String,
    target_digest: String,
}

impl Diagnostics {
    fn of(plan: &TailAppendSourcePublicationPlan) -> Self {
        Self {
            source_bytes: plan.source_bytes(),
            inserted_bytes: plan.inserted_bytes(),
            output_bytes: plan.output_bytes(),
            input_bytes: plan.input_bytes(),
            paragraphs: plan.paragraphs(),
            runs: plan.runs(),
            source_digest: plan.source_digest().to_owned(),
            target_digest: plan.target_digest().to_owned(),
        }
    }

    /// Whether the commit changed the package's output.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.inserted_bytes > 0
    }

    /// Exact source length.
    #[must_use]
    pub const fn source_bytes(&self) -> usize {
        self.source_bytes
    }

    /// Bytes inserted at the body tail.
    #[must_use]
    pub const fn inserted_bytes(&self) -> usize {
        self.inserted_bytes
    }

    /// Exact length the committed package writes.
    #[must_use]
    pub const fn output_bytes(&self) -> usize {
        self.output_bytes
    }

    /// UTF-8 bytes copied from the staged runs.
    #[must_use]
    pub const fn input_bytes(&self) -> usize {
        self.input_bytes
    }

    /// Appended paragraphs.
    #[must_use]
    pub const fn paragraphs(&self) -> usize {
        self.paragraphs
    }

    /// Appended runs.
    #[must_use]
    pub const fn runs(&self) -> usize {
        self.runs
    }

    /// Canonical lowercase SHA-256 digest of the source.
    #[must_use]
    pub fn source_digest(&self) -> &str {
        &self.source_digest
    }

    /// Canonical lowercase SHA-256 digest of the committed output.
    #[must_use]
    pub fn target_digest(&self) -> &str {
        &self.target_digest
    }
}

/// An exact, reversible source-backed tail-append patch: the pair of
/// packages it was committed between. It replays onto the recorded source
/// content, including an independently opened equal-content package, and its
/// inverse onto the recorded target content. Other content is refused.
#[derive(Clone)]
pub struct Patch {
    before: Package,
    after: Package,
    diagnostics: Diagnostics,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("is_noop", &self.is_noop())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether applying the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.same_package(&self.after)
    }

    /// The patch that undoes this one, applicable to its result.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            diagnostics: self.diagnostics.clone(),
        }
    }
}

impl Package {
    /// Start a source-backed tail append of this package's body with the
    /// default append bounds.
    ///
    /// # Errors
    ///
    /// Returns the materialization errors of a committed package's output
    /// (see the [module](super) docs); an opened package cannot fail.
    pub fn edit_tail_append(&self) -> Result<Edit, Error> {
        self.edit_tail_append_with_options(EditOptions::new())
    }

    /// [`Self::edit_tail_append`] with explicit append bounds.
    ///
    /// # Errors
    ///
    /// As [`Self::edit_tail_append`].
    pub fn edit_tail_append_with_options(&self, options: EditOptions) -> Result<Edit, Error> {
        let (base, inner) = self.staging(options)?;
        Ok(Edit { base, inner })
    }

    /// Replay a source-backed tail-append patch onto the recorded source
    /// content, including an independently opened equal-content package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Plan`] with a conflict for another package.
    pub fn apply_tail_append(&self, patch: &Patch) -> Result<Commit, Error> {
        if self.same_package(&patch.before) {
            return Ok(Commit::new(
                patch.after.clone(),
                patch.clone(),
                patch.diagnostics.clone(),
            ));
        }

        let recipient = self.replay_digest()?;
        let recorded = patch.before.replay_digest()?;
        if recipient != recorded {
            return Err(Error::Plan(crate::tail_append::Error::PatchConflict));
        }
        if !patch.diagnostics.changed() {
            return Ok(Commit::new(
                self.clone(),
                patch.clone(),
                patch.diagnostics.clone(),
            ));
        }

        let (plan, inverse) = if let Some(plan) = patch.after.pending_plan() {
            (plan, false)
        } else if let Some(plan) = patch.before.pending_plan() {
            (plan, true)
        } else {
            return Err(Error::Plan(crate::tail_append::Error::PatchConflict));
        };
        let after = self.rebase_plan(plan, inverse)?;
        Ok(Commit::new(after, patch.clone(), patch.diagnostics.clone()))
    }
}
