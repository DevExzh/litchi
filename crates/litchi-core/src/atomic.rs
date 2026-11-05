//! Format-neutral atomic filesystem publication.
//!
//! A save stages a complete artifact in a uniquely named sibling file, then
//! publishes that file with one replacement operation.  The staging callback
//! owns serialization and may return any caller error; this module owns only
//! the filesystem sequence and converts its own failures through [`Error`].

use std::error::Error as StdError;
use std::ffi::OsString;
use std::fmt;
use std::fs::{self, File, OpenOptions, Permissions};
use std::io::{self, BufWriter, ErrorKind, Write};
use std::path::{Path, PathBuf};
use std::result::Result;
use std::sync::atomic::{AtomicU64, Ordering};

use crate::Durability;

const MAX_TEMP_FILE_ATTEMPTS: usize = 128;
static TEMP_FILE_COUNTER: AtomicU64 = AtomicU64::new(0);

/// Whether publication may replace an existing regular-file destination.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Overwrite {
    /// Replace the destination atomically.
    Replace,
    /// Create the destination only if it is still absent at publication time.
    Refuse,
}

/// An error raised by the format-neutral publication sequence.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The artifact was not published because a filesystem operation failed.
    Io(io::Error),
    /// The destination was published, but a later cleanup or durability step
    /// failed. Callers must not retry this result as though publication had
    /// not happened.
    Committed {
        /// The cleanup or durability error observed after publication.
        source: io::Error,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Io(source) => write!(formatter, "atomic publication failed: {source}"),
            Self::Committed { source } => write!(
                formatter,
                "publication committed but cleanup or durability synchronization failed: {source}"
            ),
        }
    }
}

impl StdError for Error {
    fn source(&self) -> Option<&(dyn StdError + 'static)> {
        match self {
            Self::Io(source) | Self::Committed { source } => Some(source),
        }
    }
}

impl From<io::Error> for Error {
    fn from(source: io::Error) -> Self {
        Self::Io(source)
    }
}

/// Stage and atomically publish one complete artifact.
///
/// The destination is validated before staging and again immediately before
/// publication.  A failed callback, synchronization, identity check, or
/// replacement leaves the destination unchanged and removes the owned
/// temporary pathname.  [`Overwrite::Refuse`] uses a same-directory hard link
/// so a destination created during staging wins the race without being
/// replaced.  Existing regular-file permissions are copied to the published
/// artifact.
///
/// The pathname and temporary-identity checks are best-effort checks.  A
/// caller must provide a trusted, private parent directory: portable pathname
/// operations cannot close every race in a hostile shared directory.
/// If identity capture fails after a staging path has been created, the save
/// is refused before serialization and that unauthenticated path may remain
/// for the caller to remove.
pub fn save<T, E, Stage, Verify>(
    destination: &Path,
    durability: Durability,
    overwrite: Overwrite,
    stage: Stage,
    verify: Verify,
) -> Result<T, E>
where
    Stage: FnOnce(&mut BufWriter<File>) -> Result<T, E>,
    Verify: FnOnce() -> Result<(), E>,
    E: From<Error>,
{
    let mut steps = SystemSteps;
    save_with_steps(
        destination,
        durability,
        overwrite,
        &mut steps,
        stage,
        verify,
    )
}

/// The filesystem operations of one publication.  Keeping these calls behind
/// a small seam makes the pre- and post-commit error contract testable without
/// changing the production sequence.
trait PublishSteps {
    fn sync_file(&mut self, file: &File) -> io::Result<()>;
    fn replace(&mut self, temporary: &Path, destination: &Path) -> PersistResult;
    fn replace_noclobber(&mut self, temporary: &Path, destination: &Path) -> PersistResult;
    fn sync_parent(&mut self, parent: &Path) -> io::Result<()>;
}

struct SystemSteps;

impl PublishSteps for SystemSteps {
    fn sync_file(&mut self, file: &File) -> io::Result<()> {
        file.sync_all()
    }

    fn replace(&mut self, temporary: &Path, destination: &Path) -> PersistResult {
        atomic_replace(temporary, destination).map_err(PersistFailure::Before)
    }

    fn replace_noclobber(&mut self, temporary: &Path, destination: &Path) -> PersistResult {
        fs::hard_link(temporary, destination).map_err(PersistFailure::Before)?;
        fs::remove_file(temporary).map_err(PersistFailure::Committed)
    }

    fn sync_parent(&mut self, parent: &Path) -> io::Result<()> {
        sync_parent(parent)
    }
}

enum PersistFailure {
    Before(io::Error),
    Committed(io::Error),
}

type PersistResult = Result<(), PersistFailure>;

fn save_with_steps<T, E, S, Stage, Verify>(
    destination: &Path,
    durability: Durability,
    overwrite: Overwrite,
    steps: &mut S,
    stage: Stage,
    verify: Verify,
) -> Result<T, E>
where
    S: PublishSteps + ?Sized,
    Stage: FnOnce(&mut BufWriter<File>) -> Result<T, E>,
    Verify: FnOnce() -> Result<(), E>,
    E: From<Error>,
{
    ensure_supported_platform::<E>()?;
    // This first check gives callers an early, deterministic refusal for a
    // directory or symbolic link and avoids creating a temporary pathname for
    // an unusable destination.
    let _ = destination_permissions(destination).map_err(|source| E::from(Error::Io(source)))?;
    let parent = parent_directory(destination);
    let (temporary_path, temporary_file) =
        create_sibling_temp_file(destination).map_err(|source| E::from(Error::Io(source)))?;
    let identity = match file_identity(&temporary_file) {
        Ok(identity) => identity,
        Err(source) => {
            // The pathname cannot be authenticated when native identity
            // capture fails.  Leave it behind rather than risking removal of
            // a different file that took its place; callers should account
            // for a possible orphaned staging pathname in this rare case.
            drop(temporary_file);
            return Err(E::from(Error::Io(source)));
        },
    };
    let mut cleanup = TemporaryCleanupGuard::new(temporary_path, identity);

    let mut buffered = BufWriter::new(temporary_file);
    let staged = stage(&mut buffered)?;
    buffered
        .flush()
        .map_err(|source| E::from(Error::Io(source)))?;
    let file = buffered
        .into_inner()
        .map_err(|error| E::from(Error::Io(error.into_error())))?;

    // Re-read the destination after serialization.  This detects an observed
    // symlink/non-file substitution and picks up the current mode when a
    // pre-existing regular file was replaced while staging.  The final check
    // below is still required because `verify` runs after this one.
    let permissions =
        destination_permissions(destination).map_err(|source| E::from(Error::Io(source)))?;
    if let Some(permissions) = permissions {
        file.set_permissions(permissions)
            .map_err(|source| E::from(Error::Io(source)))?;
    }
    if durability.syncs_file() {
        steps
            .sync_file(&file)
            .map_err(|source| E::from(Error::Io(source)))?;
    }
    verify()?;

    // Windows requires every handle to be closed before the temporary pathname
    // can replace the destination.
    drop(file);
    ensure_temp_identity(cleanup.path(), identity).map_err(|source| E::from(Error::Io(source)))?;
    // `verify` is caller-owned and may change the destination after the
    // preceding check.  Revalidate immediately before publication.
    let _ = destination_permissions(destination).map_err(|source| E::from(Error::Io(source)))?;

    let persist = match overwrite {
        Overwrite::Replace => steps.replace(cleanup.path(), destination),
        Overwrite::Refuse => steps.replace_noclobber(cleanup.path(), destination),
    };
    match persist {
        Ok(()) => cleanup.mark_published(),
        Err(PersistFailure::Before(source)) => return Err(E::from(Error::Io(source))),
        Err(PersistFailure::Committed(source)) => {
            cleanup.mark_published();
            return Err(E::from(Error::Committed { source }));
        },
    }
    if durability.syncs_directory() {
        steps
            .sync_parent(parent)
            .map_err(|source| E::from(Error::Committed { source }))?;
    }
    Ok(staged)
}

#[cfg(any(unix, windows))]
fn ensure_supported_platform<E>() -> Result<(), E>
where
    E: From<Error>,
{
    Ok(())
}

#[cfg(not(any(unix, windows)))]
fn ensure_supported_platform<E>() -> Result<(), E>
where
    E: From<Error>,
{
    Err(E::from(Error::Io(io::Error::new(
        ErrorKind::Unsupported,
        "atomic filesystem publication requires a platform file identity",
    ))))
}

fn parent_directory(path: &Path) -> &Path {
    path.parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."))
}

fn create_sibling_temp_file(path: &Path) -> io::Result<(PathBuf, File)> {
    let file_name = path.file_name().ok_or_else(|| {
        io::Error::new(
            ErrorKind::InvalidInput,
            "atomic output path must name a file",
        )
    })?;
    let parent = parent_directory(path);

    for _ in 0..MAX_TEMP_FILE_ATTEMPTS {
        let counter = TEMP_FILE_COUNTER.fetch_add(1, Ordering::Relaxed);
        let mut temporary_name = OsString::from(".");
        temporary_name.push(file_name);
        temporary_name.push(format!(
            ".litchi-atomic-{}-{counter}.tmp",
            std::process::id()
        ));
        let temporary_path = parent.join(temporary_name);
        match OpenOptions::new()
            .read(true)
            .write(true)
            .create_new(true)
            .open(&temporary_path)
        {
            Ok(file) => return Ok((temporary_path, file)),
            Err(source) if source.kind() == ErrorKind::AlreadyExists => {},
            Err(source) => return Err(source),
        }
    }

    Err(io::Error::new(
        ErrorKind::AlreadyExists,
        "could not allocate a unique sibling temporary file",
    ))
}

fn destination_permissions(path: &Path) -> io::Result<Option<Permissions>> {
    if path.file_name().is_none() {
        return Err(io::Error::new(
            ErrorKind::InvalidInput,
            "atomic output path must name a file",
        ));
    }
    match fs::symlink_metadata(path) {
        Ok(metadata) if metadata.file_type().is_symlink() => Err(io::Error::new(
            ErrorKind::InvalidInput,
            "atomic output path must not be a symbolic link",
        )),
        Ok(metadata) if !metadata.is_file() => Err(io::Error::new(
            ErrorKind::InvalidInput,
            "atomic output path must be a regular file",
        )),
        Ok(metadata) => Ok(Some(metadata.permissions())),
        Err(source) if source.kind() == ErrorKind::NotFound => Ok(None),
        Err(source) => Err(source),
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TempIdentity {
    #[cfg(unix)]
    Unix { device: u64, inode: u64 },
    #[cfg(windows)]
    Windows { volume: u32, index: u64 },
    // `save_with_steps` refuses this target before creating a staging file;
    // this fallback only keeps the private type total for compilation.
    #[cfg(not(any(unix, windows)))]
    Fallback { length: u64 },
}

fn file_identity(file: &File) -> io::Result<TempIdentity> {
    identity_from_metadata(&file.metadata()?)
}

fn path_identity(path: &Path) -> io::Result<TempIdentity> {
    let metadata = fs::symlink_metadata(path)?;
    if !metadata.file_type().is_file() {
        return Err(io::Error::new(
            ErrorKind::InvalidData,
            "atomic temporary path is not a regular file",
        ));
    }
    identity_from_metadata(&metadata)
}

fn identity_from_metadata(metadata: &fs::Metadata) -> io::Result<TempIdentity> {
    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;

        Ok(TempIdentity::Unix {
            device: metadata.dev(),
            inode: metadata.ino(),
        })
    }

    #[cfg(windows)]
    {
        use std::os::windows::fs::MetadataExt;

        let volume = metadata.volume_serial_number().ok_or_else(|| {
            io::Error::new(
                ErrorKind::Unsupported,
                "atomic temporary file has no volume identity",
            )
        })?;
        let index = metadata.file_index().ok_or_else(|| {
            io::Error::new(
                ErrorKind::Unsupported,
                "atomic temporary file has no file identity",
            )
        })?;
        Ok(TempIdentity::Windows { volume, index })
    }

    #[cfg(not(any(unix, windows)))]
    {
        Ok(TempIdentity::Fallback {
            length: metadata.len(),
        })
    }
}

fn ensure_temp_identity(path: &Path, expected: TempIdentity) -> io::Result<()> {
    let observed = path_identity(path)?;
    if observed == expected {
        Ok(())
    } else {
        Err(io::Error::new(
            ErrorKind::AlreadyExists,
            "atomic temporary path changed while staged output was open",
        ))
    }
}

struct TemporaryCleanupGuard {
    path: PathBuf,
    identity: TempIdentity,
    published: bool,
}

impl TemporaryCleanupGuard {
    fn new(path: PathBuf, identity: TempIdentity) -> Self {
        Self {
            path,
            identity,
            published: false,
        }
    }

    fn path(&self) -> &Path {
        &self.path
    }

    fn mark_published(&mut self) {
        self.published = true;
    }
}

impl Drop for TemporaryCleanupGuard {
    fn drop(&mut self) {
        if self.published {
            return;
        }
        // Never unlink a pathname after another file has taken its place.
        // This remains best-effort and therefore requires a trusted/private
        // parent directory; a portable path check and unlink cannot close
        // every hostile-directory race.
        cleanup_owned_temp(&self.path, self.identity);
    }
}

fn cleanup_owned_temp(path: &Path, expected: TempIdentity) {
    if !path_identity(path).is_ok_and(|observed| observed == expected) {
        return;
    }

    #[cfg(windows)]
    {
        // A replacement can inherit read-only permissions from its original
        // destination.  Windows refuses to unlink that authenticated file
        // until the bit is cleared.  Recheck identity after changing the
        // metadata so a pathname substitution is still preserved.
        let Ok(metadata) = fs::metadata(path) else {
            return;
        };
        let mut permissions = metadata.permissions();
        if permissions.readonly() {
            permissions.set_readonly(false);
            if fs::set_permissions(path, permissions).is_err() {
                return;
            }
            if !path_identity(path).is_ok_and(|observed| observed == expected) {
                return;
            }
        }
    }

    drop(fs::remove_file(path));
}

#[cfg(not(windows))]
fn atomic_replace(temporary: &Path, destination: &Path) -> io::Result<()> {
    fs::rename(temporary, destination)
}

#[cfg(windows)]
fn atomic_replace(temporary: &Path, destination: &Path) -> io::Result<()> {
    use std::os::windows::ffi::OsStrExt;
    use windows_sys::Win32::Storage::FileSystem::{
        MOVEFILE_REPLACE_EXISTING, MOVEFILE_WRITE_THROUGH, MoveFileExW,
    };

    let temporary: Vec<u16> = temporary
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();
    let destination: Vec<u16> = destination
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();
    // SAFETY: both vectors are NUL-terminated UTF-16 strings that remain
    // alive for the duration of the OS call.
    let replaced = unsafe {
        MoveFileExW(
            temporary.as_ptr(),
            destination.as_ptr(),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH,
        )
    };
    if replaced == 0 {
        Err(io::Error::last_os_error())
    } else {
        Ok(())
    }
}

#[cfg(unix)]
fn sync_parent(parent: &Path) -> io::Result<()> {
    match File::open(parent).and_then(|directory| directory.sync_all()) {
        Ok(()) => Ok(()),
        Err(source)
            if matches!(
                source.kind(),
                ErrorKind::InvalidInput | ErrorKind::Unsupported
            ) =>
        {
            Ok(())
        },
        Err(source) => Err(source),
    }
}

#[cfg(not(unix))]
fn sync_parent(_parent: &Path) -> io::Result<()> {
    Ok(())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "focused publication tests use panic-on-failure assertions"
    )]

    use super::*;
    use std::fs;
    use std::io::Write;

    #[derive(Debug, Default)]
    struct CountingSteps {
        file_syncs: usize,
        replacements: usize,
        parent_syncs: usize,
        fail_file_sync: bool,
        fail_parent_sync: bool,
        fail_noclobber_cleanup: bool,
    }

    impl PublishSteps for CountingSteps {
        fn sync_file(&mut self, file: &File) -> io::Result<()> {
            self.file_syncs += 1;
            if self.fail_file_sync {
                Err(io::Error::other("injected file sync failure"))
            } else {
                file.sync_all()
            }
        }

        fn replace(&mut self, temporary: &Path, destination: &Path) -> PersistResult {
            self.replacements += 1;
            atomic_replace(temporary, destination).map_err(PersistFailure::Before)
        }

        fn replace_noclobber(&mut self, temporary: &Path, destination: &Path) -> PersistResult {
            self.replacements += 1;
            fs::hard_link(temporary, destination).map_err(PersistFailure::Before)?;
            if self.fail_noclobber_cleanup {
                Err(PersistFailure::Committed(io::Error::other(
                    "injected no-clobber cleanup failure",
                )))
            } else {
                fs::remove_file(temporary).map_err(PersistFailure::Committed)
            }
        }

        fn sync_parent(&mut self, parent: &Path) -> io::Result<()> {
            self.parent_syncs += 1;
            if self.fail_parent_sync {
                Err(io::Error::other("injected parent sync failure"))
            } else {
                sync_parent(parent)
            }
        }
    }

    fn save_test_file(
        directory: &Path,
        name: &str,
        contents: &[u8],
        durability: Durability,
        overwrite: Overwrite,
        steps: &mut impl PublishSteps,
    ) -> Result<usize, Error> {
        let destination = directory.join(name);
        save_with_steps(
            &destination,
            durability,
            overwrite,
            steps,
            |staging| {
                staging.write_all(contents)?;
                Ok(contents.len())
            },
            || Ok(()),
        )
    }

    #[test]
    fn durability_levels_only_call_their_selected_sync_steps() {
        for (durability, expected_file_syncs, expected_parent_syncs) in [
            (Durability::Full, 1, 1),
            (Durability::FileOnly, 1, 0),
            (Durability::NoSync, 0, 0),
        ] {
            let directory = tempfile::tempdir().expect("temporary directory");
            let mut steps = CountingSteps::default();
            let bytes = save_test_file(
                directory.path(),
                "document.rtf",
                b"new",
                durability,
                Overwrite::Replace,
                &mut steps,
            )
            .expect("publication");
            assert_eq!(bytes, 3);
            assert_eq!(
                fs::read(directory.path().join("document.rtf")).unwrap(),
                b"new"
            );
            assert_eq!(steps.file_syncs, expected_file_syncs);
            assert_eq!(steps.parent_syncs, expected_parent_syncs);
            assert_eq!(steps.replacements, 1);
        }
    }

    #[test]
    fn pre_commit_failure_preserves_destination_and_cleans_staging() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.rtf");
        fs::write(&destination, b"old").expect("seed destination");
        let mut steps = CountingSteps {
            fail_file_sync: true,
            ..CountingSteps::default()
        };

        let result = save_test_file(
            directory.path(),
            "document.rtf",
            b"new",
            Durability::Full,
            Overwrite::Replace,
            &mut steps,
        );
        assert!(
            matches!(result, Err(Error::Io(source)) if source.to_string().contains("injected"))
        );
        assert_eq!(fs::read(destination).unwrap(), b"old");
        assert_eq!(fs::read_dir(directory.path()).unwrap().count(), 1);
        assert_eq!(
            (steps.file_syncs, steps.replacements, steps.parent_syncs),
            (1, 0, 0)
        );
    }

    #[test]
    fn post_commit_failure_is_distinguishable_and_keeps_new_destination() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let mut steps = CountingSteps {
            fail_parent_sync: true,
            ..CountingSteps::default()
        };

        let result = save_test_file(
            directory.path(),
            "document.rtf",
            b"new",
            Durability::Full,
            Overwrite::Replace,
            &mut steps,
        );
        assert!(
            matches!(result, Err(Error::Committed { source }) if source.to_string().contains("injected"))
        );
        assert_eq!(
            fs::read(directory.path().join("document.rtf")).unwrap(),
            b"new"
        );
        assert_eq!(fs::read_dir(directory.path()).unwrap().count(), 1);
    }

    #[test]
    fn no_clobber_refuses_existing_destination_without_replacing_it() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.rtf");
        fs::write(&destination, b"old").expect("seed destination");
        let mut steps = CountingSteps::default();

        let result = save_test_file(
            directory.path(),
            "document.rtf",
            b"new",
            Durability::NoSync,
            Overwrite::Refuse,
            &mut steps,
        );
        assert!(
            matches!(result, Err(Error::Io(source)) if source.kind() == ErrorKind::AlreadyExists)
        );
        assert_eq!(fs::read(destination).unwrap(), b"old");
        assert_eq!(fs::read_dir(directory.path()).unwrap().count(), 1);
    }

    #[cfg(unix)]
    #[test]
    fn temporary_path_substitution_is_refused_without_unlinking_substitute() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.rtf");
        fs::write(&destination, b"old").expect("seed destination");
        let temporary_prefix = ".document.rtf.litchi-atomic-";

        let result: Result<(), Error> = save_with_steps(
            &destination,
            Durability::NoSync,
            Overwrite::Replace,
            &mut CountingSteps::default(),
            |staging| {
                staging.write_all(b"new")?;
                Ok(())
            },
            || {
                let temporary = fs::read_dir(&directory)
                    .expect("list staging directory")
                    .map(|entry| entry.expect("entry").path())
                    .find(|path| {
                        path.file_name()
                            .and_then(|name| name.to_str())
                            .is_some_and(|name| name.starts_with(temporary_prefix))
                    })
                    .expect("staging path");
                fs::remove_file(&temporary).expect("remove owned pathname");
                fs::write(&temporary, b"substitute").expect("replace pathname");
                Ok(())
            },
        );

        assert!(
            matches!(result, Err(Error::Io(source)) if source.kind() == ErrorKind::AlreadyExists)
        );
        assert_eq!(fs::read(destination).unwrap(), b"old");
        let substitute = fs::read_dir(&directory)
            .unwrap()
            .map(|entry| entry.unwrap().path())
            .find(|path| {
                path.file_name()
                    .and_then(|name| name.to_str())
                    .is_some_and(|name| name.starts_with(temporary_prefix))
            })
            .expect("substitute remains");
        assert_eq!(fs::read(substitute).unwrap(), b"substitute");
    }

    #[cfg(unix)]
    #[test]
    fn destination_symlink_substitution_is_refused_and_preserved() {
        use std::os::unix::fs::symlink;

        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.rtf");
        let target = directory.path().join("attacker-target");
        fs::write(&destination, b"old").expect("seed destination");
        fs::write(&target, b"target").expect("seed target");
        let mut steps = CountingSteps::default();

        let result: Result<(), Error> = save_with_steps(
            &destination,
            Durability::NoSync,
            Overwrite::Replace,
            &mut steps,
            |staging| {
                staging.write_all(b"new")?;
                Ok(())
            },
            || {
                fs::remove_file(&destination).expect("remove original destination");
                symlink(&target, &destination).expect("substitute destination symlink");
                Ok(())
            },
        );

        assert!(matches!(
            result,
            Err(Error::Io(source)) if source.kind() == ErrorKind::InvalidInput
        ));
        assert!(
            fs::symlink_metadata(&destination)
                .expect("substituted destination")
                .file_type()
                .is_symlink()
        );
        assert_eq!(fs::read_link(destination).unwrap(), target);
        assert_eq!(
            fs::read(directory.path().join("attacker-target")).unwrap(),
            b"target"
        );
        assert_eq!(fs::read_dir(directory.path()).unwrap().count(), 2);
    }

    #[test]
    fn no_clobber_closes_the_create_race_after_the_last_destination_check() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.rtf");
        let mut steps = CountingSteps::default();
        let result: Result<usize, Error> = save_with_steps(
            &destination,
            Durability::NoSync,
            Overwrite::Refuse,
            &mut steps,
            |staging| {
                staging.write_all(b"new")?;
                Ok(3)
            },
            || {
                fs::write(&destination, b"racer").expect("create racing destination");
                Ok(())
            },
        );

        assert!(
            matches!(result, Err(Error::Io(source)) if source.kind() == ErrorKind::AlreadyExists)
        );
        assert_eq!(fs::read(destination).unwrap(), b"racer");
        assert_eq!(fs::read_dir(directory.path()).unwrap().count(), 1);
    }

    #[test]
    fn no_clobber_cleanup_failure_reports_committed_and_keeps_destination() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let mut steps = CountingSteps {
            fail_noclobber_cleanup: true,
            ..CountingSteps::default()
        };

        let result = save_test_file(
            directory.path(),
            "document.rtf",
            b"new",
            Durability::NoSync,
            Overwrite::Refuse,
            &mut steps,
        );

        assert!(matches!(
            result,
            Err(Error::Committed { source })
                if source.to_string().contains("injected no-clobber cleanup failure")
        ));
        assert_eq!(
            fs::read(directory.path().join("document.rtf")).unwrap(),
            b"new"
        );
        // The injected cleanup failure deliberately leaves the private
        // staging pathname for its owner to inspect or clean up.
        assert_eq!(fs::read_dir(directory.path()).unwrap().count(), 2);
    }

    #[test]
    fn error_keeps_the_caller_error_type() {
        #[derive(Debug, PartialEq, Eq)]
        enum CallerError {
            Stage,
            Atomic,
        }

        impl From<Error> for CallerError {
            fn from(error: Error) -> Self {
                assert!(matches!(error, Error::Io(_)));
                Self::Atomic
            }
        }

        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.rtf");
        let result: Result<(), CallerError> = save(
            &destination,
            Durability::NoSync,
            Overwrite::Replace,
            |_staging| Err(CallerError::Stage),
            || Ok(()),
        );
        assert_eq!(result, Err(CallerError::Stage));
        assert_eq!(fs::read_dir(directory.path()).unwrap().count(), 0);
    }

    #[test]
    fn verify_error_preserves_destination_and_cleans_staging() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.rtf");
        fs::write(&destination, b"old").expect("seed destination");
        let mut steps = CountingSteps::default();

        let result: Result<(), Error> = save_with_steps(
            &destination,
            Durability::NoSync,
            Overwrite::Replace,
            &mut steps,
            |staging| {
                staging.write_all(b"new")?;
                Ok(())
            },
            || Err(Error::Io(io::Error::other("injected verify failure"))),
        );

        assert!(matches!(
            result,
            Err(Error::Io(source)) if source.to_string().contains("injected verify failure")
        ));
        assert_eq!(fs::read(destination).unwrap(), b"old");
        assert_eq!(fs::read_dir(directory.path()).unwrap().count(), 1);
    }

    #[cfg(windows)]
    #[test]
    fn read_only_owned_temporary_is_cleaned_after_precommit_failure() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.rtf");
        fs::write(&destination, b"old").expect("seed destination");
        let mut destination_permissions = fs::metadata(&destination)
            .expect("destination metadata")
            .permissions();
        destination_permissions.set_readonly(true);
        fs::set_permissions(&destination, destination_permissions)
            .expect("make destination readonly");
        let mut steps = CountingSteps::default();

        let result: Result<(), Error> = save_with_steps(
            &destination,
            Durability::NoSync,
            Overwrite::Replace,
            &mut steps,
            |staging| {
                staging.write_all(b"new")?;
                Ok(())
            },
            || Err(Error::Io(io::Error::other("injected verify failure"))),
        );

        assert!(
            matches!(result, Err(Error::Io(source)) if source.to_string().contains("injected verify failure"))
        );
        assert_eq!(fs::read(&destination).unwrap(), b"old");
        assert!(fs::metadata(destination).unwrap().permissions().readonly());
        assert_eq!(fs::read_dir(directory.path()).unwrap().count(), 1);
    }
}
