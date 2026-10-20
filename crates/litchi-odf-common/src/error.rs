//! The ingress/egress error of the packaged ODF owners.
//!
//! Package-level verbs (`open*`, `from_*`, `save*`, `to_bytes`, `write_to`,
//! `Edit::commit` and `Package::apply`) of every ODF family return this type.
//! Limit failures are typed with their dimension; every other failure keeps
//! the established common-core error as its [`Error::Core`] payload, so
//! `?` conversions in either direction stay lossless for limit dimensions.

use std::sync::Arc;

use litchi_core::{Resource, ResourceLimit};

use crate::limits::{LimitKind, SourceLimitKind};

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
                    });
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
    /// A source-backed `content.xml` publication failure as an owner error:
    /// a source change and an I/O, allocation or core failure keep their
    /// typed core payload; every other publication failure keeps its message.
    fn from(error: crate::core::SourceContentPublicationError) -> Self {
        use crate::core::SourceContentPublicationError as Publication;
        match error {
            Publication::Core(error) => Self::from(error),
            Publication::SourceChanged {
                expected, observed, ..
            } => Self::Core(litchi_core::Error::SourceChanged { expected, observed }),
            Publication::Sink { source, .. } => Self::Core(litchi_core::Error::Io(source)),
            Publication::Allocation { resource, source } => {
                Self::Core(litchi_core::Error::Allocation { resource, source })
            },
            Publication::Source { source, .. } => Self::from(source),
            other => Self::Core(litchi_core::Error::Other(other.to_string())),
        }
    }
}

impl From<crate::core::source_publication::SourceContentInsertionError> for Error {
    /// A source-backed `content.xml` insertion failure as an owner error; a
    /// publication failure keeps the mapping of its publication error.
    fn from(error: crate::core::source_publication::SourceContentInsertionError) -> Self {
        use crate::core::source_publication::SourceContentInsertionError as Insertion;
        match error {
            Insertion::Publication { source, .. } => Self::from(source),
            other => Self::Core(litchi_core::Error::Other(other.to_string())),
        }
    }
}

impl From<Error> for litchi_core::Error {
    fn from(error: Error) -> Self {
        match error {
            Error::Core(error) => error,
            Error::Limit {
                kind,
                observed,
                maximum,
            } => Self::ResourceLimit(ResourceLimit {
                resource: match kind {
                    LimitKind::Source(SourceLimitKind::Members) => Resource::Objects,
                    LimitKind::Semantic(kind) => kind.resource(),
                    _ => Resource::InputBytes,
                },
                observed,
                limit: maximum,
                scope: Arc::from(match kind {
                    LimitKind::Source(kind) => kind.scope(),
                    LimitKind::Semantic(kind) => kind.scope(),
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
        }
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
}
