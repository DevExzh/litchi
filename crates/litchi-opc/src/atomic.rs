//! Atomic filesystem replacement for finalized package artifacts.

use std::fs::{self, File, Permissions};
use std::io;
use std::path::Path;

use crate::error::{OpcError, Result};

/// Finalize an artifact in a sibling temporary file, then replace `path`.
///
/// The destination is left untouched when `write` fails. Existing regular-file
/// permissions are preserved on Unix, and symbolic-link or non-file
/// destinations are rejected before any replacement is attempted. Windows
/// permission preservation is not currently promised.
///
/// # Errors
///
/// Returns an error when `path` is not a usable file destination, when the
/// temporary file cannot be written, synchronized, or persisted, or when the
/// `write` callback fails. If the destination was already replaced but the
/// parent directory could not be synchronized, the error is
/// [`OpcError::Committed`].
pub fn replace(path: &Path, write: impl FnOnce(&mut File) -> Result<()>) -> Result<()> {
    replace_with(path, write)
}

/// Atomically replace `path` while preserving a caller-owned typed error.
///
/// Filesystem validation and persistence failures are converted from
/// [`OpcError`], while an error returned by `write` is passed through exactly.
///
/// # Errors
///
/// Returns an error when `path` is not a usable file destination, when the
/// temporary file cannot be written, synchronized, or persisted, or when the
/// `write` callback fails. Filesystem failures are converted into `E` from
/// [`OpcError`]; a `write` failure is returned unchanged.
pub fn replace_with<E>(
    path: &Path,
    write: impl FnOnce(&mut File) -> std::result::Result<(), E>,
) -> std::result::Result<(), E>
where
    E: From<OpcError>,
{
    replace_with_impl(path, write, sync_parent)
}

fn replace_with_impl<E, S>(
    path: &Path,
    write: impl FnOnce(&mut File) -> std::result::Result<(), E>,
    sync: S,
) -> std::result::Result<(), E>
where
    E: From<OpcError>,
    S: FnOnce(&Path) -> io::Result<()>,
{
    if path.file_name().is_none() {
        return Err(E::from(invalid_path(
            "package destination must name a file",
        )));
    }

    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let permissions = destination_permissions(path).map_err(E::from)?;
    let mut temporary = tempfile::Builder::new()
        .prefix(".litchi-")
        .suffix(".tmp")
        .tempfile_in(parent)
        .map_err(OpcError::from)
        .map_err(E::from)?;

    write(temporary.as_file_mut())?;
    if let Some(existing_permissions) = permissions {
        temporary
            .as_file()
            .set_permissions(existing_permissions)
            .map_err(OpcError::from)
            .map_err(E::from)?;
    }
    temporary
        .as_file_mut()
        .sync_all()
        .map_err(OpcError::from)
        .map_err(E::from)?;
    let _persisted = temporary
        .persist(path)
        .map_err(|error| OpcError::IoError(error.error))
        .map_err(E::from)?;
    sync(parent).map_err(|source| E::from(OpcError::Committed { source }))?;
    Ok(())
}

fn destination_permissions(path: &Path) -> Result<Option<Permissions>> {
    match fs::symlink_metadata(path) {
        Ok(metadata) if metadata.file_type().is_symlink() => Err(invalid_path(
            "refusing to replace a package destination through a symbolic link",
        )),
        Ok(metadata) if !metadata.is_file() => {
            Err(invalid_path("package destination is not a regular file"))
        },
        Ok(metadata) => Ok(Some(metadata.permissions())),
        Err(error) if error.kind() == io::ErrorKind::NotFound => Ok(None),
        Err(error) => Err(error.into()),
    }
}

#[cfg(unix)]
fn sync_parent(parent: &Path) -> io::Result<()> {
    match File::open(parent).and_then(|directory| directory.sync_all()) {
        Ok(()) => Ok(()),
        Err(error)
            if matches!(
                error.kind(),
                io::ErrorKind::InvalidInput | io::ErrorKind::Unsupported
            ) =>
        {
            Ok(())
        },
        Err(error) => Err(error),
    }
}

#[cfg(not(unix))]
fn sync_parent(_parent: &Path) -> io::Result<()> {
    Ok(())
}

fn invalid_path(message: &'static str) -> OpcError {
    OpcError::IoError(io::Error::new(io::ErrorKind::InvalidInput, message))
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use std::io::Write;

    use super::*;

    #[derive(Debug)]
    enum TypedError {
        Opc(OpcError),
        Write,
    }

    impl From<OpcError> for TypedError {
        fn from(error: OpcError) -> Self {
            Self::Opc(error)
        }
    }

    impl std::fmt::Display for TypedError {
        fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
            match self {
                Self::Opc(error) => write!(formatter, "{error}"),
                Self::Write => formatter.write_str("typed write failure"),
            }
        }
    }

    impl std::error::Error for TypedError {}

    #[test]
    fn failed_finalization_leaves_the_destination_untouched() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("report.xlsx");
        fs::write(&destination, b"original").expect("seed destination");

        let result = replace(&destination, |temporary| {
            temporary.write_all(b"partial")?;
            Err(OpcError::InvalidRelationship(
                "injected finalization failure".to_owned(),
            ))
        });

        assert!(matches!(result, Err(OpcError::InvalidRelationship(_))));
        assert_eq!(
            fs::read(&destination).expect("read destination"),
            b"original"
        );
        assert_eq!(
            fs::read_dir(directory.path())
                .expect("list temporary directory")
                .count(),
            1
        );
    }

    #[test]
    fn successful_finalization_replaces_the_destination() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("report.xlsx");
        fs::write(&destination, b"old").expect("seed destination");

        replace(&destination, |temporary| {
            temporary.write_all(b"new")?;
            Ok(())
        })
        .expect("atomic replacement");

        assert_eq!(fs::read(destination).expect("read destination"), b"new");
    }

    #[cfg(unix)]
    #[test]
    fn directory_sync_failure_reports_that_replacement_already_committed() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("report.xlsx");
        fs::write(&destination, b"old").expect("seed destination");

        let result = replace_with_impl::<TypedError, _>(
            &destination,
            |temporary| {
                temporary
                    .write_all(b"new")
                    .map_err(|error| TypedError::Opc(OpcError::IoError(error)))?;
                Ok(())
            },
            |_parent| {
                Err(io::Error::new(
                    io::ErrorKind::PermissionDenied,
                    "injected directory sync failure",
                ))
            },
        );

        assert!(matches!(
            result,
            Err(TypedError::Opc(OpcError::Committed { source }))
                if source.kind() == io::ErrorKind::PermissionDenied
        ));
        assert_eq!(fs::read(destination).expect("read destination"), b"new");
    }

    #[test]
    fn caller_owned_write_error_is_not_erased() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("report.docx");
        fs::write(&destination, b"original").expect("seed destination");

        let result = replace_with::<TypedError>(&destination, |_temporary| Err(TypedError::Write));

        assert!(matches!(result, Err(TypedError::Write)));
        assert_eq!(
            fs::read(destination).expect("read destination"),
            b"original"
        );
    }

    #[cfg(unix)]
    #[test]
    fn replacement_preserves_existing_file_permissions() {
        use std::os::unix::fs::PermissionsExt;

        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("report.xlsx");
        fs::write(&destination, b"old").expect("seed destination");
        fs::set_permissions(&destination, Permissions::from_mode(0o640))
            .expect("set destination permissions");

        replace(&destination, |temporary| {
            temporary.write_all(b"new")?;
            Ok(())
        })
        .expect("atomic replacement");

        let mode = fs::metadata(destination)
            .expect("destination metadata")
            .permissions()
            .mode()
            & 0o777;
        assert_eq!(mode, 0o640);
    }

    #[cfg(unix)]
    #[test]
    fn symbolic_link_destinations_are_refused_without_touching_the_target() {
        use std::os::unix::fs::symlink;

        let directory = tempfile::tempdir().expect("temporary directory");
        let target = directory.path().join("target.xlsx");
        let link = directory.path().join("report.xlsx");
        fs::write(&target, b"original").expect("seed target");
        symlink(&target, &link).expect("create symbolic link");

        let result = replace(&link, |temporary| {
            temporary.write_all(b"replacement")?;
            Ok(())
        });

        assert!(matches!(
            result,
            Err(OpcError::IoError(error)) if error.kind() == io::ErrorKind::InvalidInput
        ));
        assert_eq!(fs::read(target).expect("read target"), b"original");
        assert!(
            fs::symlink_metadata(link)
                .expect("link metadata")
                .file_type()
                .is_symlink()
        );
    }
}
