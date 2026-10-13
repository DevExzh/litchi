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
            Self::InvalidLimit { .. } | Self::Limit { .. } => None,
        }
    }

    /// Return the limit dimension of a limit failure.
    #[must_use]
    pub const fn limit_kind(&self) -> Option<LimitKind> {
        match self {
            Self::InvalidLimit { kind, .. } | Self::Limit { kind, .. } => Some(*kind),
            Self::Core(_) => None,
        }
    }
}

impl From<litchi_core::Error> for Error {
    fn from(error: litchi_core::Error) -> Self {
        match error {
            litchi_core::Error::ResourceLimit(limit) => {
                match SourceLimitKind::from_scope(&limit.scope) {
                    Some(kind) => Self::Limit {
                        kind: LimitKind::Source(kind),
                        observed: limit.observed,
                        maximum: limit.limit,
                    },
                    None => Self::Core(litchi_core::Error::ResourceLimit(limit)),
                }
            },
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
                    LimitKind::Semantic(crate::limits::SemanticLimitKind::Units) => {
                        Resource::Objects
                    },
                    _ => Resource::InputBytes,
                },
                observed,
                limit: maximum,
                scope: Arc::from(match kind {
                    LimitKind::Source(kind) => kind.scope(),
                    LimitKind::Semantic(kind) => kind.as_str(),
                }),
            }),
            invalid @ Error::InvalidLimit { .. } => Self::InvalidFormat(invalid.to_string()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::Error;
    use crate::limits::{LimitKind, SourceLimitKind};

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
