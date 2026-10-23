//! The story-hyperlink redaction family of a source-backed package
//! (record 0355a): remove the hyperlinks to selected target URLs from every
//! relationship-owned story (main, headers, footers, notes, comments,
//! glossary), with their relationship records.
//!
//! The redaction is forward-only and sealed to the exact source it was
//! planned on; its inverse restores the exact source on that commit's
//! retained owner, while a detached equal artifact remains refused.

use std::sync::Arc;

use crate::Result;
use crate::source_backed::Package;
use crate::source_backed::deferred::{Change, source_family};
use crate::story_hyperlinks::{self as record, EffectReport, Mode, Report};

/// The sealed forward redaction, or its inverse.
#[derive(Clone)]
pub(crate) enum StoryRedactionChange {
    Forward(Arc<record::ForwardOnlyPatch>),
    Inverse {
        noop: bool,
        forward: Arc<record::ForwardOnlyPatch>,
    },
}

impl Change for StoryRedactionChange {
    type Error = crate::Error;

    fn is_noop(&self) -> bool {
        match self {
            Self::Forward(patch) => patch.effect_report().is_noop(),
            Self::Inverse { noop, .. } => *noop,
        }
    }

    fn allows_foreign_restore(&self) -> bool {
        match self {
            Self::Forward(_) => true,
            Self::Inverse { noop, .. } => *noop,
        }
    }

    fn inverse(&self) -> Self {
        match self {
            Self::Forward(patch) => Self::Inverse {
                noop: patch.effect_report().is_noop(),
                forward: Arc::clone(patch),
            },
            Self::Inverse { forward, .. } => Self::Forward(Arc::clone(forward)),
        }
    }

    fn publish(&self, package: &Package) -> Result<Package> {
        match self {
            Self::Forward(patch) => {
                package.story_hyperlink_redaction_admit(patch)?;
                if patch.effect_report().is_noop() {
                    return Ok(package.clone());
                }
                let base = package.clone();
                let stream_base = base.clone();
                let ordinary_patch = Arc::clone(patch);
                let stream_patch = Arc::clone(patch);
                package.deferred(
                    move |writer, policy| {
                        base.write_story_hyperlink_redaction(writer, &ordinary_patch, policy)
                    },
                    move |target| {
                        let parts = record::publication_parts(&stream_patch)?;
                        stream_base.stream_relationship_overlays(target, parts)
                    },
                )
            },
            Self::Inverse { noop: true, .. } => Ok(package.clone()),
            Self::Inverse { .. } => Err(crate::Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "story_hyperlink_redaction",
                "forward-only redaction cannot be inverted on another package",
            )),
        }
    }
}

source_family!("story-hyperlink redaction", StoryRedactionChange);

impl From<record::ForwardOnlyPatch> for Patch {
    /// Replay a sealed record commit onto the exact source it was planned on.
    fn from(patch: record::ForwardOnlyPatch) -> Self {
        Self {
            change: StoryRedactionChange::Forward(Arc::new(patch)),
            restore: None,
        }
    }
}

/// The budgets of this family, reported by
/// [`crate::Error::ExternalHyperlinkRedactionLimit`].
pub use crate::redact::LimitKind;

/// A staged story-hyperlink redaction. The source is never changed:
/// [`Self::commit`] returns a new package.
#[derive(Clone)]
pub struct Edit {
    package: Package,
    plan: record::Plan,
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

    /// The predicted effects, without changing anything.
    #[must_use]
    pub const fn effect_report(&self) -> EffectReport {
        self.plan.effect_report()
    }

    /// The completeness report, including diagnostics.
    #[must_use]
    pub const fn report(&self) -> &Report {
        self.plan.report()
    }

    /// Commit the redaction as a new source-backed package.
    ///
    /// # Errors
    ///
    /// Returns the redaction, source, protection or signature errors.
    pub fn commit(self) -> Result<Commit> {
        let patch = self.plan.apply()?;
        commit_change(
            &self.package,
            StoryRedactionChange::Forward(Arc::new(patch)),
        )
    }
}

impl Package {
    /// Stage the removal of the hyperlinks to `target_urls` from every
    /// relationship-owned story.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::plan_story_hyperlink_redaction`].
    pub fn edit_story_hyperlink_redaction(&self, target_urls: &[&str], mode: Mode) -> Result<Edit> {
        Ok(Edit {
            package: self.clone(),
            plan: self.plan_story_hyperlink_redaction(target_urls, mode)?,
        })
    }

    /// Replay a story-hyperlink redaction patch onto this package.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for another source, or a typed
    /// refusal for a non-noop inverse applied outside its retained owner.
    pub fn apply_story_hyperlink_redaction(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}
