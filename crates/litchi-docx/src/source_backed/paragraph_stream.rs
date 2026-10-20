//! The authored paragraph-stream append family of a source-backed package
//! (record 0355a): append paragraphs from a replayable source or a one-shot
//! producer to the main story, never materializing the main document.
//!
//! [`Package::tail_append_plain_paragraphs`](crate::source_backed::Package::tail_append_plain_paragraphs)
//! stages an [`Edit`]; its commit returns a new source-backed package whose
//! write streams the append from the source and the authored replay.

use std::sync::Arc;

use crate::source_backed::Package;
use crate::source_backed::deferred::{Change, source_family};
use crate::source_backed::tail_append_stream::{
    self as stream, AuthoredInput, ParagraphStreamLimits, ParagraphStreamPlan, PatchExpectations,
    PreparedStream,
};

pub use crate::source_backed::tail_append_stream::ParagraphStreamEdit as Edit;

/// The prepared append (with what replaying it needs), or its restore-only
/// inverse.
#[derive(Clone)]
pub(crate) enum StreamChange {
    Forward {
        input: Arc<dyn AuthoredInput>,
        limits: ParagraphStreamLimits,
        expected: Option<PatchExpectations>,
        noop: bool,
        prepared: Option<Arc<PreparedStream>>,
    },
    Inverse {
        noop: bool,
        forward: Box<StreamChange>,
    },
}

impl Change for StreamChange {
    type Error = stream::Error;

    fn is_noop(&self) -> bool {
        match self {
            Self::Forward { noop, .. } | Self::Inverse { noop, .. } => *noop,
        }
    }

    fn inverse(&self) -> Self {
        match self {
            Self::Forward { noop, .. } => Self::Inverse {
                noop: *noop,
                forward: Box::new(self.clone()),
            },
            Self::Inverse { forward, .. } => (**forward).clone(),
        }
    }

    fn publish(&self, package: &Package) -> Result<Package, Error> {
        match self {
            Self::Forward {
                input,
                limits,
                expected,
                prepared,
                ..
            } => {
                let prepared = match prepared {
                    Some(prepared) if prepared.base.same_owner(package) => Arc::clone(prepared),
                    _ => Arc::new(
                        package
                            .stream_edit(Arc::clone(input), *limits, expected.clone())
                            .prepare()?
                            .into_prepared(),
                    ),
                };
                if prepared.is_noop() {
                    return Ok(package.clone());
                }
                Ok(package.deferred(move |writer| {
                    prepared.write(writer).map_err(|error| match error {
                        stream::Error::TailAppend(
                            crate::source_backed::tail_append::Error::Document(error),
                        ) => error,
                        stream::Error::Opc(error) => crate::Error::Opc(error),
                        other => crate::Error::Other(other.to_string()),
                    })
                })?)
            },
            Self::Inverse { noop: true, .. } => Ok(package.clone()),
            Self::Inverse { .. } => Err(Error::Invalid(
                "a paragraph-stream inverse restores only its commit's own package",
            )),
        }
    }
}

source_family!("paragraph-stream append", StreamChange, stream::Error);

/// The budgets of this family, reported by [`Error::Limit`].
pub use crate::source_backed::tail_append_stream::LimitKind;

pub(crate) fn commit_plan(plan: ParagraphStreamPlan<'_>) -> Result<Commit, Error> {
    let package = plan.package_share();
    let (input, expected) = plan.replay_input();
    let limits = plan.limits();
    let prepared = Arc::new(plan.into_prepared());
    let change = StreamChange::Forward {
        input,
        limits,
        expected,
        noop: prepared.is_noop(),
        prepared: Some(prepared),
    };
    commit_change(&package, change)
}

impl Package {
    /// Replay a paragraph-stream patch onto this package: the same authored
    /// paragraphs are appended (a one-shot producer cannot be replayed).
    ///
    /// # Errors
    ///
    /// Returns the preparation errors, or an inverse applied elsewhere than
    /// its commit's own package.
    pub fn apply_paragraph_stream(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_change(self, patch)
    }
}
