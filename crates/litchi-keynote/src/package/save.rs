//! Durable filesystem publication for an immutable Keynote package.

use std::fmt;
use std::path::Path;

use litchi_iwa_archive::publication::{Error as PublicationError, replace_with};

use super::{Package, WriteError};

/// A content-redacted failure while durably saving a Keynote package.
///
/// Writing the package to a private staging file and publishing that file are
/// separate failure domains. The distinction is retained so callers can tell
/// whether the destination was potentially replaced. The contained errors are
/// available through accessors and remain the source of this error for callers
/// that need to inspect their typed details.
/// Default `Debug` and `Display` output is content-redacted; explicitly
/// traversing [`std::error::Error::source`] or the accessors may reveal the
/// caller-owned sink or operating-system diagnostic.
#[non_exhaustive]
pub enum SaveError {
    /// The package could not be written to its staging file.
    Write(WriteError),
    /// The staged package could not be published to the destination.
    Publication(PublicationError),
}

impl SaveError {
    /// Borrow the package write failure, if this is a write error.
    #[must_use]
    pub const fn write_error(&self) -> Option<&WriteError> {
        match self {
            Self::Write(error) => Some(error),
            Self::Publication(_) => None,
        }
    }

    /// Borrow the filesystem publication failure, if this is a publication
    /// error.
    #[must_use]
    pub const fn publication_error(&self) -> Option<&PublicationError> {
        match self {
            Self::Write(_) => None,
            Self::Publication(error) => Some(error),
        }
    }

    /// Return whether the destination replacement may already have committed.
    ///
    /// A write failure always occurs before publication. A publication failure
    /// delegates to the archive publication boundary's committed-state bit.
    #[must_use]
    pub fn was_committed(&self) -> bool {
        match self {
            Self::Write(_) => false,
            Self::Publication(error) => error.was_committed(),
        }
    }

    /// Consume this error and return its package write failure, if any.
    #[must_use]
    pub fn into_write_error(self) -> Option<WriteError> {
        match self {
            Self::Write(error) => Some(error),
            Self::Publication(_) => None,
        }
    }

    /// Consume this error and return its filesystem publication failure, if
    /// any.
    #[must_use]
    pub fn into_publication_error(self) -> Option<PublicationError> {
        match self {
            Self::Write(_) => None,
            Self::Publication(error) => Some(error),
        }
    }
}

impl From<WriteError> for SaveError {
    fn from(error: WriteError) -> Self {
        Self::Write(error)
    }
}

impl From<PublicationError> for SaveError {
    fn from(error: PublicationError) -> Self {
        Self::Publication(error)
    }
}

impl fmt::Debug for SaveError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let mut debug = formatter.debug_struct("SaveError");
        match self {
            Self::Write(error) => {
                debug
                    .field("kind", &"write")
                    .field("bytes_written", &error.bytes_written());
            },
            Self::Publication(error) => {
                debug
                    .field("kind", &"publication")
                    .field("was_committed", &error.was_committed());
            },
        }
        debug.finish_non_exhaustive()
    }
}

impl fmt::Display for SaveError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Write(_) => "could not write Keynote package while staging it",
            Self::Publication(_) => "could not publish Keynote package",
        })
    }
}

impl std::error::Error for SaveError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Write(error) => Some(error),
            Self::Publication(error) => Some(error),
        }
    }
}

/// Save one package through the archive layer's durable publication boundary.
pub(super) fn save(package: &Package, path: impl AsRef<Path>) -> Result<(), SaveError> {
    replace_with::<SaveError>(path.as_ref(), |temporary| {
        package.write_to(temporary).map_err(SaveError::Write)
    })
}
