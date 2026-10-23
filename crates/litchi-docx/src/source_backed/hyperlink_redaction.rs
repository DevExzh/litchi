//! The irreversible external-hyperlink redaction family of a
//! source-backed package (record 0355a): unwrap the main document's
//! hyperlinks to selected target URLs and remove their relationship
//! records.
//!
//! The redaction is forward-only: its record patch keeps no removed
//! content, so its inverse restores only the exact source of the commit it
//! came from ([`Package::apply_external_hyperlink_redaction`] on that
//! commit's package).
//!
//! [`Package::apply_external_hyperlink_redaction`]: crate::source_backed::Package::apply_external_hyperlink_redaction

use std::sync::Arc;

use crate::Result;
use crate::redact::{self as record, EffectReport};
use crate::source_backed::Package;
use crate::source_backed::deferred::{Change, source_family};

/// The sealed forward redaction, or its inverse.
#[derive(Clone)]
pub(crate) enum RedactionChange {
    Forward(Arc<record::Commit>),
    Inverse {
        noop: bool,
        forward: Arc<record::Commit>,
    },
}

impl Change for RedactionChange {
    type Error = crate::Error;

    fn is_noop(&self) -> bool {
        match self {
            Self::Forward(commit) => commit.patch().is_noop(),
            Self::Inverse { noop, .. } => *noop,
        }
    }

    fn inverse(&self) -> Self {
        match self {
            Self::Forward(commit) => Self::Inverse {
                noop: commit.patch().is_noop(),
                forward: Arc::clone(commit),
            },
            Self::Inverse { forward, .. } => Self::Forward(Arc::clone(forward)),
        }
    }

    fn publish(&self, package: &Package) -> Result<Package> {
        match self {
            Self::Forward(commit) => {
                let main = package.external_hyperlink_redaction_target(commit)?;
                if commit.patch().is_noop() {
                    return Ok(package.clone());
                }
                let base = package.clone();
                let stream_base = base.clone();
                let ordinary_commit = Arc::clone(commit);
                let stream_commit = Arc::clone(commit);
                let stream_main = main.clone();
                package.deferred(
                    move |writer, policy| {
                        base.write_external_hyperlink_redaction(
                            writer,
                            &main,
                            &ordinary_commit,
                            policy,
                        )
                    },
                    move |target| {
                        let (replacement, removed_ids) = redact::publication_parts(&stream_commit);
                        stream_base.stream_relationship_overlays(
                            target,
                            [(stream_main.clone(), replacement, removed_ids)],
                        )
                    },
                )
            },
            Self::Inverse { noop: true, .. } => Ok(package.clone()),
            // Only the commit's own package can be restored.
            Self::Inverse { .. } => Err(crate::Error::ExternalHyperlinkRedactionConflict),
        }
    }
}

source_family!("external-hyperlink redaction", RedactionChange);

impl From<record::Commit> for Patch {
    /// Replay a sealed record commit onto the exact source it was planned on.
    fn from(commit: record::Commit) -> Self {
        Self {
            change: RedactionChange::Forward(Arc::new(commit)),
            restore: None,
        }
    }
}

/// The budgets of this family, reported by
/// [`crate::Error::ExternalHyperlinkRedactionLimit`].
pub use crate::redact::LimitKind;

/// A staged redaction of the main document's hyperlinks to selected
/// target URLs. The source is never changed: [`Self::commit`] returns a new
/// package.
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

    /// Commit the redaction as a new source-backed package.
    ///
    /// # Errors
    ///
    /// Returns the redaction, source, protection or signature errors.
    pub fn commit(self) -> Result<Commit> {
        let commit = self.plan.apply()?;
        commit_change(&self.package, RedactionChange::Forward(Arc::new(commit)))
    }
}

impl Package {
    /// Stage the redaction of the main document's hyperlinks to
    /// `target_urls`.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::plan_external_hyperlink_redaction`].
    pub fn edit_external_hyperlink_redaction(&self, target_urls: &[&str]) -> Result<Edit> {
        Ok(Edit {
            package: self.clone(),
            plan: self.plan_external_hyperlink_redaction(target_urls)?,
        })
    }

    /// Replay a redaction patch onto this package.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::ExternalHyperlinkRedactionConflict`] for another
    /// source, or for an inverse applied to anything but its commit's
    /// package.
    pub fn apply_external_hyperlink_redaction(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}
