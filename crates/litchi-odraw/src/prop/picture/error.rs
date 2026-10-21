use std::{fmt, io};

pub use litchi_core::edit::{Refusal, RefusalKind};

/// Resources bounded by picture property admission and publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum LimitKind {
    /// Bytes in the property record body, excluding the wire header.
    PropertyBytes,
    /// Six-byte property descriptors in one table.
    Properties,
    /// UTF-16 code units in the picture name, excluding the terminator.
    NameUnits,
}

/// Picture package failures preserve structural, policy and replay distinctions.
#[non_exhaustive]
pub enum Error {
    Format(crate::Error),
    Limit {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    InvalidLimit {
        kind: LimitKind,
        value: u64,
        maximum: u64,
    },
    Refused(Refusal),
    PatchConflict,
    Io(io::Error),
}

impl From<crate::Error> for Error {
    fn from(error: crate::Error) -> Self {
        Self::Format(error)
    }
}

impl From<io::Error> for Error {
    fn from(error: io::Error) -> Self {
        Self::Io(error)
    }
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Format(error) => error.fmt(f),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                f,
                "picture {kind:?} budget exceeded: {observed} > {maximum}"
            ),
            Self::InvalidLimit {
                kind,
                value,
                maximum,
            } => write!(f, "invalid picture {kind:?} limit: {value} > {maximum}"),
            Self::Refused(refusal) => refusal.fmt(f),
            Self::PatchConflict => f.write_str("picture patch base does not match recipient"),
            Self::Io(error) => write!(f, "picture I/O failure ({:?})", error.kind()),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Format(error) => Some(error),
            Self::Refused(refusal) => Some(refusal),
            Self::Io(error) => Some(error),
            _ => None,
        }
    }
}

pub type Result<T> = std::result::Result<T, Error>;

pub(super) fn check(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(Error::Limit {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        });
    }
    Ok(())
}

impl fmt::Debug for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Format(error) => f.debug_tuple("Format").field(error).finish(),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => f
                .debug_struct("Limit")
                .field("kind", kind)
                .field("observed", observed)
                .field("maximum", maximum)
                .finish(),
            Self::InvalidLimit {
                kind,
                value,
                maximum,
            } => f
                .debug_struct("InvalidLimit")
                .field("kind", kind)
                .field("value", value)
                .field("maximum", maximum)
                .finish(),
            Self::Refused(refusal) => f.debug_tuple("Refused").field(refusal).finish(),
            Self::PatchConflict => f.write_str("PatchConflict"),
            Self::Io(error) => f
                .debug_struct("Io")
                .field("kind", &error.kind())
                .finish_non_exhaustive(),
        }
    }
}
