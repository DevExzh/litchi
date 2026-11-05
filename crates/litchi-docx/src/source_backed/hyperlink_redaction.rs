//! The irreversible external-hyperlink redaction family of a
//! source-backed package (record 0355a): unwrap the main document's
//! hyperlinks to selected target URLs and remove their relationship
//! records.
//!
//! The redaction is forward-only: its record patch keeps no removed
//! content, so its inverse restores the exact source on the commit's
//! retained owner. A detached equal artifact still refuses that inverse.
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
    Forward(Arc<record::IrreversibleRedactionPatch>),
    Inverse {
        noop: bool,
        forward: Arc<record::IrreversibleRedactionPatch>,
    },
}

impl Change for RedactionChange {
    type Error = crate::Error;

    fn is_noop(&self) -> bool {
        match self {
            Self::Forward(patch) => patch.is_noop(),
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
                noop: patch.is_noop(),
                forward: Arc::clone(patch),
            },
            Self::Inverse { forward, .. } => Self::Forward(Arc::clone(forward)),
        }
    }

    fn publish(&self, package: &Package) -> Result<Package> {
        match self {
            Self::Forward(patch) => {
                let main = package.external_hyperlink_redaction_target(patch)?;
                if patch.is_noop() {
                    return Ok(package.clone());
                }
                let base = package.clone();
                let stream_base = base.clone();
                let ordinary_patch = Arc::clone(patch);
                let stream_patch = Arc::clone(patch);
                let stream_main = main.clone();
                package.deferred(
                    move |writer, policy| {
                        base.write_external_hyperlink_redaction(
                            writer,
                            &main,
                            &ordinary_patch,
                            policy,
                        )
                    },
                    move |target| {
                        let (replacement, removed_ids) = record::publication_parts(&stream_patch);
                        stream_base.stream_relationship_overlays(
                            target,
                            [(stream_main.clone(), replacement, removed_ids)],
                        )
                    },
                )
            },
            Self::Inverse { noop: true, .. } => Ok(package.clone()),
            // Only the commit's own package can be restored.
            Self::Inverse { .. } => Err(crate::Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "external_hyperlink_redaction",
                "forward-only redaction cannot be inverted on another package",
            )),
        }
    }
}

source_family!("external-hyperlink redaction", RedactionChange);

impl From<record::IrreversibleRedactionPatch> for Patch {
    /// Replay a sealed record commit onto the exact source it was planned on.
    fn from(commit: record::IrreversibleRedactionPatch) -> Self {
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
        let patch = self.plan.apply()?;
        commit_change(&self.package, RedactionChange::Forward(Arc::new(patch)))
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
    /// Returns [`crate::Error::PatchConflict`] for another source, or a typed
    /// refusal for a non-noop inverse applied outside its retained owner.
    pub fn apply_external_hyperlink_redaction(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}
