//! Atomic filesystem replacement for finalized package artifacts.

use std::fs::{self, File, Permissions};
use std::io;
use std::path::Path;

use crate::error::{OpcError, Result};

pub(crate) fn replace(path: &Path, write: impl FnOnce(&mut File) -> Result<()>) -> Result<()> {
    if path.file_name().is_none() {
        return Err(invalid_path("package destination must name a file"));
    }

    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let permissions = destination_permissions(path)?;
    let mut temporary = tempfile::Builder::new()
        .prefix(".litchi-")
        .suffix(".tmp")
        .tempfile_in(parent)?;

    write(temporary.as_file_mut())?;
    if let Some(permissions) = permissions {
        temporary.as_file().set_permissions(permissions)?;
    }
    temporary.as_file_mut().sync_all()?;
    let _persisted = temporary
        .persist(path)
        .map_err(|error| OpcError::IoError(error.error))?;
    sync_parent(parent)?;
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
fn sync_parent(parent: &Path) -> Result<()> {
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
        Err(error) => Err(error.into()),
    }
}

#[cfg(not(unix))]
fn sync_parent(_parent: &Path) -> Result<()> {
    Ok(())
}

fn invalid_path(message: &'static str) -> OpcError {
    OpcError::IoError(io::Error::new(io::ErrorKind::InvalidInput, message))
}

#[cfg(test)]
mod tests {
    use std::io::Write;

    use super::*;

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
