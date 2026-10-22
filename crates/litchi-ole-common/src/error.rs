//! Typed failures at bounded shared OLE boundaries.

use std::{error, fmt, io};

use litchi_cfb::OleError;
use litchi_core::edit::{Refusal, RefusalKind};

/// Finite resource dimensions enforced by shared OLE readers and editors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum LimitKind {
    /// Selected object count.
    Objects,
    /// Storage path depth.
    StorageDepth,
    /// Streams in one selected object.
    ObjectStreams,
    /// Streams retained in the complete package.
    PackageStreams,
    /// Bytes in one stream.
    StreamBytes,
    /// Bytes in one rendered object.
    ObjectBytes,
    /// Aggregate retained stream bytes.
    TotalBytes,
    /// Descendant storages in one selected object.
    ObjectStorages,
    /// Storages retained in the complete package.
    PackageStorages,
    /// Directory entries visited under an object capture budget.
    DirectoryEntries,
    /// UTF-8 bytes retained for one raw directory name or CLSID.
    DirectoryNameBytes,
    /// Aggregate UTF-8 bytes retained by a raw directory catalog.
    DirectoryTotalBytes,
    /// Recursive depth visited by a raw directory catalog.
    DirectoryDepth,
    /// Stream removals in one operation.
    StreamRemovals,
    /// Components in a stream selector.
    SelectorDepth,
    /// Encoded bytes in one user-defined hyperlink BLOB.
    UserDefinedBlobBytes,
    /// Hyperlink records in one user-defined `_PID_HLINKS` value.
    UserDefinedLinks,
    /// UTF-16 code units in one user-defined hyperlink string.
    UserDefinedStringUnits,
    /// Aggregate UTF-16 code units in one user-defined hyperlink value.
    UserDefinedTotalUtf16Units,
}

/// Shared OLE errors preserve low-level CFB failures and identify typed budgets.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// A low-level CFB or raw metadata failure.
    Cfb(OleError),
    /// A finite object resource budget was exceeded.
    Limit {
        /// Resource dimension.
        kind: LimitKind,
        /// Observed resource usage.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// An object limit is outside its supported nonzero range.
    InvalidLimit {
        /// Resource dimension.
        kind: LimitKind,
        /// Requested value.
        value: u64,
        /// Largest supported value.
        maximum: u64,
    },
    /// The recipient does not match the patch's recorded base.
    PatchConflict,
    /// A well-formed object request that this route refuses.
    ///
    /// The operation and reason are static, so selectors and admission
    /// decisions never echo caller-supplied object names into an error.
    Refused(Refusal),
}

impl From<OleError> for Error {
    #[cold]
    #[inline(never)]
    fn from(error: OleError) -> Self {
        match error {
            OleError::LimitExceeded {
                resource,
                observed,
                maximum,
            } => match limit_kind(resource) {
                Some(kind) => Self::Limit {
                    kind,
                    observed,
                    maximum,
                },
                None => Self::Cfb(OleError::LimitExceeded {
                    resource,
                    observed,
                    maximum,
                }),
            },
            OleError::InvalidLimit {
                resource,
                value,
                maximum,
            } => match limit_kind(resource) {
                Some(kind) => Self::InvalidLimit {
                    kind,
                    value,
                    maximum,
                },
                None => Self::Cfb(OleError::InvalidLimit {
                    resource,
                    value,
                    maximum,
                }),
            },
            error => Self::Cfb(error),
        }
    }
}

/// Maps only labels owned by this common object/raw layer. Physical CFB
/// labels and labels introduced by another helper stay in [`Error::Cfb`]
/// rather than being assigned a misleading object dimension.
fn limit_kind(resource: &str) -> Option<LimitKind> {
    Some(match resource {
        "directory entries" | "protected directory entries" => LimitKind::DirectoryEntries,
        "directory name bytes" => LimitKind::DirectoryNameBytes,
        "directory total bytes" => LimitKind::DirectoryTotalBytes,
        "directory depth" => LimitKind::DirectoryDepth,
        "link stream bytes" => LimitKind::StreamBytes,
        "stream removal selectors" => LimitKind::StreamRemovals,
        _ => return None,
    })
}

impl From<io::Error> for Error {
    fn from(error: io::Error) -> Self {
        Self::Cfb(OleError::Io(error))
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Cfb(error) => error.fmt(formatter),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => {
                write!(
                    formatter,
                    "OLE {kind:?} budget exceeded: {observed} > {maximum}"
                )
            },
            Self::InvalidLimit {
                kind,
                value,
                maximum,
            } => {
                write!(
                    formatter,
                    "invalid OLE {kind:?} limit: {value}; supported range 1..={maximum}"
                )
            },
            Self::PatchConflict => formatter.write_str("OLE patch base does not match recipient"),
            Self::Refused(refusal) => refusal.fmt(formatter),
        }
    }
}

impl error::Error for Error {
    fn source(&self) -> Option<&(dyn error::Error + 'static)> {
        match self {
            Self::Cfb(error) => Some(error),
            Self::Refused(refusal) => Some(refusal),
            _ => None,
        }
    }
}

impl LimitKind {
    fn resource(self) -> litchi_core::Resource {
        match self {
            Self::StorageDepth | Self::DirectoryDepth | Self::SelectorDepth => {
                litchi_core::Resource::Depth
            },
            Self::StreamBytes | Self::ObjectBytes | Self::TotalBytes => {
                litchi_core::Resource::InputBytes
            },
            Self::DirectoryNameBytes | Self::DirectoryTotalBytes => {
                litchi_core::Resource::InputBytes
            },
            Self::UserDefinedBlobBytes
            | Self::UserDefinedStringUnits
            | Self::UserDefinedTotalUtf16Units => litchi_core::Resource::InputBytes,
            Self::Objects
            | Self::ObjectStreams
            | Self::PackageStreams
            | Self::ObjectStorages
            | Self::PackageStorages
            | Self::DirectoryEntries
            | Self::StreamRemovals
            | Self::UserDefinedLinks => litchi_core::Resource::Objects,
        }
    }
}

impl From<Error> for litchi_core::Error {
    fn from(error: Error) -> Self {
        match error {
            Error::Cfb(error @ OleError::SourceChanged { .. }) => Self::Io(io::Error::other(error)),
            Error::Cfb(OleError::InvalidLimit { value, maximum, .. }) => Self::InvalidLimit {
                kind: litchi_core::Resource::InputBytes,
                value,
                maximum,
                scope: "OLE container",
            },
            Error::Cfb(error) => error.into(),
            Error::Limit {
                kind,
                observed,
                maximum,
            } => Self::ResourceLimit(litchi_core::ResourceLimit {
                resource: kind.resource(),
                observed,
                limit: maximum,
                scope: format!("OLE {kind:?}").into(),
            }),
            Error::InvalidLimit {
                kind,
                value,
                maximum,
            } => Self::InvalidLimit {
                kind: kind.resource(),
                value,
                maximum,
                scope: "OLE object",
            },
            Error::PatchConflict => Self::PatchConflict,
            Error::Refused(refusal) => Self::Refused(refusal),
        }
    }
}

impl Error {
    /// A content-free refusal of `operation` for `reason`.
    #[must_use]
    pub const fn refused(kind: RefusalKind, operation: &'static str, reason: &'static str) -> Self {
        Self::Refused(Refusal::new(kind, operation, reason))
    }

    /// The selected object target is absent from the admitted catalog.
    #[must_use]
    pub(crate) const fn target_not_found() -> Self {
        Self::refused(
            RefusalKind::NotFound,
            "object_target",
            "the selected object target does not exist",
        )
    }

    /// A content-free refusal for a caller-supplied object selector.
    #[must_use]
    pub(crate) const fn invalid_selector() -> Self {
        Self::refused(
            RefusalKind::InvalidValue,
            "object_selector",
            "the object selector is invalid",
        )
    }

    /// A content-free refusal for a caller-supplied Property Set selector.
    #[must_use]
    pub(crate) const fn invalid_property_set_selector() -> Self {
        Self::refused(
            RefusalKind::InvalidValue,
            "property_set_selector",
            "the Property Set selector is invalid",
        )
    }

    /// A content-free refusal for an absent Property Set element.
    #[must_use]
    pub(crate) const fn property_set_element_not_found() -> Self {
        Self::refused(
            RefusalKind::NotFound,
            "property_set_element",
            "the Property Set element does not exist",
        )
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, LimitKind};
    use litchi_cfb::OleError;

    #[test]
    fn raw_limit_labels_keep_their_exact_dimensions() {
        let cases = [
            ("directory entries", LimitKind::DirectoryEntries),
            ("protected directory entries", LimitKind::DirectoryEntries),
            ("directory name bytes", LimitKind::DirectoryNameBytes),
            ("directory total bytes", LimitKind::DirectoryTotalBytes),
            ("directory depth", LimitKind::DirectoryDepth),
            ("link stream bytes", LimitKind::StreamBytes),
            ("stream removal selectors", LimitKind::StreamRemovals),
        ];
        for (resource, kind) in cases {
            let error = Error::from(OleError::LimitExceeded {
                resource,
                observed: 2,
                maximum: 1,
            });
            assert!(matches!(
                error,
                Error::Limit {
                    kind: actual,
                    observed: 2,
                    maximum: 1,
                } if actual == kind
            ));

            let error = Error::from(OleError::InvalidLimit {
                resource,
                value: 0,
                maximum: 1,
            });
            assert!(matches!(
                error,
                Error::InvalidLimit {
                    kind: actual,
                    value: 0,
                    maximum: 1,
                } if actual == kind
            ));
        }
    }

    #[test]
    fn unknown_cfb_limit_labels_stay_low_level() {
        let error = Error::from(OleError::LimitExceeded {
            resource: "CFB input bytes",
            observed: 2,
            maximum: 1,
        });
        assert!(matches!(
            error,
            Error::Cfb(OleError::LimitExceeded {
                resource: "CFB input bytes",
                observed: 2,
                maximum: 1,
            })
        ));

        let error = Error::from(OleError::InvalidLimit {
            resource: "CFB allocation table bytes",
            value: 0,
            maximum: 1,
        });
        assert!(matches!(
            error,
            Error::Cfb(OleError::InvalidLimit {
                resource: "CFB allocation table bytes",
                value: 0,
                maximum: 1,
            })
        ));
    }
}
