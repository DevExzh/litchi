//! The ingress/egress error of the packaged ODF owners.
//!
//! Package-level verbs (`open*`, `from_*`, `save*`, `to_bytes`, `write_to`,
//! `Edit::commit` and `Package::apply`) of every ODF family return this type.
//! Limit failures are typed with their dimension; every other failure keeps
//! the established common-core error as its [`Error::Core`] payload, while
//! source-backed publication failures retain their typed progress and
//! diagnostics. The ODF-facing error retains those details; conversion to the
//! common-core error intentionally projects only the classification needed by
//! core callers, while limit dimensions remain typed at the ODF boundary.

use std::sync::Arc;

use litchi_core::{Resource, ResourceLimit};

use crate::limits::{LimitKind, RawXmlLimitKind, SourceLimitKind};

/// An ODF package ingress, egress or publication failure.
///
/// Messages and `Debug` output carry dimension names, sizes and structural
/// metadata only, never authored content.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A limit profile was constructed with a value outside `1..=maximum`.
    #[error("invalid ODF {kind} limit {value} (allowed 1..={maximum})")]
    InvalidLimit {
        /// The dimension being configured.
        kind: LimitKind,
        /// The rejected value.
        value: u64,
        /// The hard ceiling for this dimension.
        maximum: u64,
    },
    /// An ingress or projection dimension exceeded the selected limit.
    #[error("ODF {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exceeded dimension.
        kind: LimitKind,
        /// The observed (or first excessive) value.
        observed: u64,
        /// The selected maximum.
        maximum: u64,
    },
    /// A well-formed request that this route refuses, classified by its
    /// [`litchi_core::edit::RefusalKind`]. The refusal carries static
    /// operation and reason strings only.
    #[error("ODF edit refused: {0}")]
    Refused(litchi_core::edit::Refusal),
    /// A patch or staged plan was made against another package base.
    #[error("ODF patch conflict in {operation}: {reason}")]
    PatchConflict {
        /// The staging or replay operation that detected the conflict.
        operation: &'static str,
        /// Why the base does not match.
        reason: &'static str,
    },
    /// A changed write was requested on a digitally signed package; an exact
    /// no-op commit keeps the signatures byte-identical instead.
    #[error("ODF package is signed: a changed write would invalidate its signatures")]
    Signed,
    /// A source-backed content publication failed before insertion replay.
    ///
    /// The box keeps the ordinary package error layout bounded; allocation is
    /// paid only on this error path so source progress and typed causes remain
    /// available to `save` callers.
    #[error(transparent)]
    SourceContentPublication(Box<crate::core::SourceContentPublicationError>),
    /// A source-backed content insertion failed during ZIP replay or its
    /// preflight. The typed replay retains source-member and secondary
    /// callback diagnostics.
    #[error(transparent)]
    SourceContentInsertion(Box<crate::core::source_publication::SourceContentInsertionError>),
    /// Any other package, parse, I/O or publication failure.
    #[error(transparent)]
    Core(litchi_core::Error),
}

impl Error {
    /// Return the common-core payload of an [`Error::Core`] failure.
    #[must_use]
    pub const fn core(&self) -> Option<&litchi_core::Error> {
        match self {
            Self::Core(error) => Some(error),
            _ => None,
        }
    }

    /// Return the limit dimension of a limit failure.
    #[must_use]
    pub const fn limit_kind(&self) -> Option<LimitKind> {
        match self {
            Self::InvalidLimit { kind, .. } | Self::Limit { kind, .. } => Some(*kind),
            _ => None,
        }
    }

    /// Return the refusal of an [`Error::Refused`] failure.
    #[must_use]
    pub const fn refusal(&self) -> Option<&litchi_core::edit::Refusal> {
        match self {
            Self::Refused(refusal) => Some(refusal),
            _ => None,
        }
    }

    /// Build an [`Error::Refused`] failure.
    #[must_use]
    pub const fn refused(
        kind: litchi_core::edit::RefusalKind,
        operation: &'static str,
        reason: &'static str,
    ) -> Self {
        Self::Refused(litchi_core::edit::Refusal::new(kind, operation, reason))
    }

    /// Build an [`Error::PatchConflict`] failure.
    #[must_use]
    pub const fn patch_conflict(operation: &'static str, reason: &'static str) -> Self {
        Self::PatchConflict { operation, reason }
    }
}

impl From<litchi_core::Error> for Error {
    fn from(error: litchi_core::Error) -> Self {
        match error {
            litchi_core::Error::ResourceLimit(limit) => {
                let kind = SourceLimitKind::from_scope(&limit.scope)
                    .map(LimitKind::Source)
                    .or_else(|| {
                        crate::limits::SemanticLimitKind::from_scope(&limit.scope)
                            .map(LimitKind::Semantic)
                    })
                    .or_else(|| RawXmlLimitKind::from_scope(&limit.scope).map(LimitKind::RawXml));
                match kind {
                    Some(kind) => Self::Limit {
                        kind,
                        observed: limit.observed,
                        maximum: limit.limit,
                    },
                    None => Self::Core(litchi_core::Error::ResourceLimit(limit)),
                }
            },
            litchi_core::Error::Refused(refusal) => Self::Refused(refusal),
            other => Self::Core(other),
        }
    }
}

impl From<std::io::Error> for Error {
    fn from(error: std::io::Error) -> Self {
        Self::Core(litchi_core::Error::Io(error))
    }
}

impl From<litchi_core::atomic::Error> for Error {
    fn from(error: litchi_core::atomic::Error) -> Self {
        Self::Core(match error {
            litchi_core::atomic::Error::Io(source) => litchi_core::Error::Io(source),
            litchi_core::atomic::Error::Committed { source } => {
                litchi_core::Error::Committed(source)
            },
            other => litchi_core::Error::Other(other.to_string()),
        })
    }
}

impl From<crate::core::SourceContentPublicationError> for Error {
    /// Preserve a source-backed `content.xml` publication failure, including
    /// its exact sink progress and typed source cause.
    fn from(error: crate::core::SourceContentPublicationError) -> Self {
        Self::SourceContentPublication(Box::new(error))
    }
}

impl From<crate::core::source_publication::SourceContentInsertionError> for Error {
    /// Preserve a source-backed `content.xml` insertion failure, including
    /// replay progress and callback diagnostics.
    fn from(error: crate::core::source_publication::SourceContentInsertionError) -> Self {
        Self::SourceContentInsertion(Box::new(error))
    }
}

impl From<Error> for litchi_core::Error {
    fn from(error: Error) -> Self {
        match error {
            Error::Limit {
                kind,
                observed,
                maximum,
            } => Self::ResourceLimit(ResourceLimit {
                resource: match kind {
                    LimitKind::Source(SourceLimitKind::Members) => Resource::Objects,
                    LimitKind::Semantic(kind) => kind.resource(),
                    LimitKind::RawXml(RawXmlLimitKind::Events)
                    | LimitKind::RawXml(RawXmlLimitKind::Attributes)
                    | LimitKind::RawXml(RawXmlLimitKind::NamespaceDeclarations)
                    | LimitKind::RawXml(RawXmlLimitKind::NamespaceBindings) => Resource::Objects,
                    LimitKind::RawXml(RawXmlLimitKind::Depth)
                    | LimitKind::RawXml(RawXmlLimitKind::NamespaceScopeDepth) => Resource::Depth,
                    LimitKind::RawXml(RawXmlLimitKind::TokenBytes)
                    | LimitKind::RawXml(RawXmlLimitKind::NamespaceBytes)
                    | LimitKind::RawXml(RawXmlLimitKind::NamespacePrefixBytes)
                    | LimitKind::RawXml(RawXmlLimitKind::NamespaceUriBytes) => Resource::Memory,
                    _ => Resource::InputBytes,
                },
                observed,
                limit: maximum,
                scope: Arc::from(match kind {
                    LimitKind::Source(kind) => kind.scope(),
                    LimitKind::Semantic(kind) => kind.scope(),
                    LimitKind::RawXml(kind) => kind.scope(),
                }),
            }),
            invalid @ Error::InvalidLimit { .. } => Self::InvalidFormat(invalid.to_string()),
            Error::Refused(refusal) => Self::Refused(refusal),
            conflict @ Error::PatchConflict { .. } => Self::InvalidFormat(conflict.to_string()),
            Error::Signed => Self::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::Protected,
                "odf.package",
                "the package is signed: a changed write would invalidate its signatures",
            )),
            Error::SourceContentPublication(error) => source_publication_core(*error),
            Error::SourceContentInsertion(error) => source_insertion_core(*error),
            Error::Core(error) => error,
        }
    }
}

fn source_publication_core(
    error: crate::core::SourceContentPublicationError,
) -> litchi_core::Error {
    use crate::core::SourceContentPublicationError as Publication;

    match error {
        Publication::Core(error) => error,
        Publication::SourceChanged {
            expected, observed, ..
        } => litchi_core::Error::SourceChanged { expected, observed },
        Publication::Sink { source, .. } => litchi_core::Error::Io(source),
        Publication::Allocation { resource, source } => {
            litchi_core::Error::Allocation { resource, source }
        },
        Publication::Source { source, .. } => source,
        other => litchi_core::Error::Other(other.to_string()),
    }
}

fn source_insertion_core(
    error: crate::core::source_publication::SourceContentInsertionError,
) -> litchi_core::Error {
    use crate::core::source_publication::SourceContentInsertionError as Insertion;

    match error {
        Insertion::Publication { source, .. } => source_publication_core(source),
        Insertion::Replay { source, .. } => replay_core(source),
        Insertion::Transport {
            publication,
            replay,
            ..
        } => {
            let publication = source_publication_core(publication);
            if matches!(publication, litchi_core::Error::SourceChanged { .. }) {
                publication
            } else {
                prefer_source_changed(publication, replay_core(*replay))
            }
        },
    }
}

fn replay_core(
    error: soapberry_zip::ReplayPublicationError<
        crate::core::source_publication::SourceContentInsertionCallbackError,
    >,
) -> litchi_core::Error {
    use soapberry_zip::ReplayPublicationError;

    match error {
        ReplayPublicationError::Archive {
            source,
            callback_error,
            ..
        } => {
            let archive = archive_core(source);
            match callback_error {
                Some(callback) => prefer_source_changed(archive, callback_core(callback)),
                None => archive,
            }
        },
        ReplayPublicationError::Callback { source, .. } => callback_core(source),
        ReplayPublicationError::Limit { .. } => {
            litchi_core::Error::Other("source-content insertion replay limit exceeded".to_string())
        },
        ReplayPublicationError::Sink {
            source,
            callback_error,
            ..
        } => {
            let sink = litchi_core::Error::Io(source);
            match callback_error {
                Some(callback) => prefer_source_changed(sink, callback_core(callback)),
                None => sink,
            }
        },
        ReplayPublicationError::NonDeterministic { .. } => litchi_core::Error::Other(
            "source-content insertion replay was non-deterministic".to_string(),
        ),
        _ => litchi_core::Error::Other("unknown source-content insertion replay failure".into()),
    }
}

fn callback_core(
    error: crate::core::source_publication::SourceContentInsertionCallbackError,
) -> litchi_core::Error {
    use crate::core::SourceMemberReaderError;
    use crate::core::source_publication::SourceContentInsertionCallbackError as Callback;

    match error {
        Callback::Member(error) => match *error {
            SourceMemberReaderError::Core { error, .. } => error,
            SourceMemberReaderError::Callback(error) => callback_core(error),
        },
        Callback::Io(error) => litchi_core::Error::Io(error),
        Callback::Cancelled => {
            litchi_core::Error::Other("source-content insertion cancelled".into())
        },
        Callback::Execution(error) => litchi_core::Error::Other(error.to_string()),
        Callback::Verification { what } => litchi_core::Error::Other(format!(
            "source-content insertion replay verification failed: {what}"
        )),
    }
}

fn archive_core(error: soapberry_zip::Error) -> litchi_core::Error {
    use soapberry_zip::ErrorKind;

    match error.into_kind() {
        ErrorKind::IO(error) | ErrorKind::CentralDirectorySpool { source: error, .. } => {
            litchi_core::Error::Io(error)
        },
        ErrorKind::FileNotFound(path) => litchi_core::Error::ComponentNotFound(path),
        ErrorKind::UnsupportedPreservation { reason } => {
            litchi_core::Error::Unsupported(reason.to_string())
        },
        kind => litchi_core::Error::ZipError(kind.to_string()),
    }
}

fn prefer_source_changed(
    primary: litchi_core::Error,
    secondary: litchi_core::Error,
) -> litchi_core::Error {
    if matches!(secondary, litchi_core::Error::SourceChanged { .. }) {
        secondary
    } else {
        primary
    }
}

#[cfg(test)]
mod tests {
    use super::Error;
    use crate::limits::{LimitKind, SourceLimitKind};

    #[test]
    fn semantic_resource_limits_become_typed_and_round_trip() {
        use crate::limits::{SemanticLimitKind, SemanticLimits};

        let limits = SemanticLimits::builder()
            .max_depth(3)
            .unwrap()
            .build()
            .unwrap();
        let core = limits.check_core(SemanticLimitKind::Depth, 4).unwrap_err();
        let typed = Error::from(core);
        assert!(matches!(
            typed,
            Error::Limit {
                kind: LimitKind::Semantic(SemanticLimitKind::Depth),
                observed: 4,
                maximum: 3,
            }
        ));
        let back = litchi_core::Error::from(typed);
        assert!(matches!(
            Error::from(back),
            Error::Limit {
                kind: LimitKind::Semantic(SemanticLimitKind::Depth),
                ..
            }
        ));
        assert!(limits.check_core(SemanticLimitKind::Depth, 3).is_ok());
    }

    #[test]
    fn refusals_conflicts_and_signatures_stay_content_free() {
        let refused = Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedTarget,
            "operation",
            "reason",
        );
        assert_eq!(
            refused.refusal().map(litchi_core::edit::Refusal::kind),
            Some(litchi_core::edit::RefusalKind::UnsupportedTarget)
        );
        assert!(
            Error::patch_conflict("operation", "reason")
                .core()
                .is_none()
        );
        assert!(matches!(
            litchi_core::Error::from(Error::Signed),
            litchi_core::Error::Unsupported(_) | litchi_core::Error::Refused(_)
        ));
    }

    #[test]
    fn source_resource_limits_become_typed_and_round_trip() {
        let core = litchi_core::Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: litchi_core::Resource::InputBytes,
            observed: 10,
            limit: 9,
            scope: std::sync::Arc::from("ODF manifest bytes"),
        });
        let typed = Error::from(core);
        assert!(matches!(
            typed,
            Error::Limit {
                kind: LimitKind::Source(SourceLimitKind::ManifestBytes),
                observed: 10,
                maximum: 9,
            }
        ));
        let flattened = litchi_core::Error::from(typed);
        assert!(matches!(
            Error::from(flattened),
            Error::Limit {
                kind: LimitKind::Source(SourceLimitKind::ManifestBytes),
                ..
            }
        ));
    }

    #[test]
    fn unrelated_resource_limits_stay_core_errors() {
        let core = litchi_core::Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: litchi_core::Resource::Depth,
            observed: 10,
            limit: 9,
            scope: std::sync::Arc::from("ODF XML depth"),
        });
        assert!(matches!(Error::from(core), Error::Core(_)));
    }

    #[test]
    fn source_insertion_conversion_retains_progress_and_source_change() {
        use crate::core::source_publication::SourceContentInsertionError;
        use crate::core::{SourceContentPublicationError, SourceContentPublicationProgress};
        use litchi_core::SourceVersion;

        let expected = SourceVersion::new(0x100, 1);
        let observed = SourceVersion::new(0x100, 2);
        let typed = Error::from(SourceContentInsertionError::Publication {
            progress: SourceContentPublicationProgress::Complete { bytes: 17 },
            source: SourceContentPublicationError::SourceChanged {
                expected,
                observed,
                progress: SourceContentPublicationProgress::Complete { bytes: 17 },
            },
        });
        let Error::SourceContentInsertion(error) = typed else {
            panic!("source insertion must remain typed");
        };
        assert_eq!(error.progress().accepted(), 17);
        assert!(matches!(
            error.as_ref(),
            SourceContentInsertionError::Publication {
                source: SourceContentPublicationError::SourceChanged {
                    expected: actual_expected,
                    observed: actual_observed,
                    progress: SourceContentPublicationProgress::Complete { bytes: 17 },
                },
                ..
            } if *actual_expected == expected && *actual_observed == observed
        ));

        let core = litchi_core::Error::from(Error::SourceContentInsertion(error));
        assert!(matches!(
            core,
            litchi_core::Error::SourceChanged {
                expected: actual_expected,
                observed: actual_observed,
            } if actual_expected == expected && actual_observed == observed
        ));
    }
}
