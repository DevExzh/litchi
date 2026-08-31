//! Durable, failure-atomic publication of one archive-owned file.
//!
//! A publication is written to a private temporary file in the destination's
//! parent directory.  The temporary file is flushed and synchronized before
//! [`tempfile::NamedTempFile::persist`] performs the platform's same-directory
//! replacement operation. Local filesystems with ordinary rename semantics
//! provide atomic name replacement; network, userspace, and unusual
//! filesystems may provide weaker guarantees.
//! On Unix, the parent directory is synchronized after the replacement when
//! the platform supports directory synchronization.  Windows uses the
//! platform replacement supplied by `tempfile`; directory synchronization is
//! unavailable there, so the file and rename sequence provide the best
//! supported durability rather than a strict directory-durability promise.
//!
//! Default `Debug` and `Display` output for the public errors intentionally
//! contains no path, operating-system error text, callback value, or archive
//! content.  The original operating-system error remains available through
//! [`Error::io_error`] and [`std::error::Error::source`] for callers that need
//! programmatic diagnostics; callers can use the typed destination and stage
//! categories for stable handling without logging those sensitive details.
//! Durable publication is supported on Unix and Windows. Other targets fail
//! with `Unsupported` before invoking the writer because this boundary cannot
//! authenticate the staging pathname there.

use std::fmt;
use std::fs::{self, File, Permissions};
use std::io::{self, Write};
use std::path::{Path, PathBuf};

/// The destination shape rejected by an atomic publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum DestinationKind {
    /// The supplied path does not name a file.
    MissingFilename,
    /// The supplied path is a symbolic link.
    Symlink,
    /// The supplied path is a Windows reparse point.
    ReparsePoint,
    /// The supplied path exists but is not a regular file.
    NonRegular,
}

impl fmt::Display for DestinationKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::MissingFilename => "missing filename",
            Self::Symlink => "symbolic link",
            Self::ReparsePoint => "reparse point",
            Self::NonRegular => "non-regular file",
        })
    }
}

/// The content-free phase at which a publication operation failed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Stage {
    /// Initial destination shape and metadata validation.
    ValidateDestination,
    /// Resolution of the destination's existing parent directory.
    ResolveParent,
    /// Creation of the same-directory private temporary file.
    CreateTemporary,
    /// Enforcement of private staging-file permissions.
    SecureTemporary,
    /// Preservation of the destination's existing permissions after
    /// replacement.
    PreservePermissions,
    /// The callback's final flush of the temporary file.
    Flush,
    /// Synchronization of the staged or published file's contents and
    /// metadata.
    SyncFile,
    /// Revalidation that the private temporary pathname still names the open
    /// staging file.
    RevalidateTemporary,
    /// Revalidation of the destination immediately before replacement.
    RevalidateDestination,
    /// Atomic replacement of the destination name.
    Persist,
    /// Synchronization of the containing directory after replacement.
    SyncParent,
}

impl fmt::Display for Stage {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::ValidateDestination => "destination validation",
            Self::ResolveParent => "destination-parent resolution",
            Self::CreateTemporary => "temporary-file creation",
            Self::SecureTemporary => "temporary-file permission hardening",
            Self::PreservePermissions => "permission preservation",
            Self::Flush => "temporary-file flush",
            Self::SyncFile => "file synchronization",
            Self::RevalidateTemporary => "temporary-file revalidation",
            Self::RevalidateDestination => "destination revalidation",
            Self::Persist => "atomic replacement",
            Self::SyncParent => "parent-directory synchronization",
        })
    }
}

/// A redacted failure from archive-owned atomic publication.
pub struct Error {
    failure: Failure,
}

/// Private representation of one publication failure.
///
/// Keeping this enum private prevents callers from constructing or
/// destructuring the retained operating-system error.  The public [`Error`]
/// type exposes only stable, content-free classifiers and explicit source
/// accessors.
enum Failure {
    /// The destination path is absent or has an unsafe/non-file shape.
    InvalidDestination {
        /// Content-free destination shape that was rejected.
        kind: DestinationKind,
        /// Publication phase in which the destination shape was observed.
        stage: Stage,
    },
    /// A filesystem operation failed before replacement committed.
    Io {
        /// Content-free publication phase that failed.
        stage: Stage,
        /// Original operating-system error, retained for programmatic
        /// diagnostics but intentionally omitted from `Debug` and `Display`.
        source: io::Error,
        /// Cached kind of the source, retained independently so callers
        /// need not inspect platform-specific error text.
        kind: io::ErrorKind,
    },
    /// The destination was replaced, but post-replacement permission or
    /// durability work failed. Callers must treat the new destination as
    /// committed.
    Committed {
        /// Content-free post-replacement publication phase that failed.
        stage: Stage,
        /// Original post-replacement error, retained for programmatic
        /// diagnostics but intentionally omitted from `Debug` and `Display`.
        source: io::Error,
        /// Cached kind of the source.
        kind: io::ErrorKind,
    },
}

impl Error {
    /// Construct an invalid-destination error without retaining path details.
    #[must_use]
    fn invalid_destination(kind: DestinationKind, stage: Stage) -> Self {
        Self {
            failure: Failure::InvalidDestination { kind, stage },
        }
    }

    /// Construct a pre-commit filesystem error without retaining OS details.
    #[must_use]
    fn io(stage: Stage, source: io::Error) -> Self {
        let kind = source.kind();
        Self {
            failure: Failure::Io {
                stage,
                source,
                kind,
            },
        }
    }

    /// Construct a committed error while retaining its post-replacement
    /// source.
    fn committed(stage: Stage, source: io::Error) -> Self {
        let kind = source.kind();
        Self {
            failure: Failure::Committed {
                stage,
                source,
                kind,
            },
        }
    }

    /// Return the rejected destination shape, when this is a validation
    /// failure.
    #[must_use]
    pub const fn destination_kind(&self) -> Option<DestinationKind> {
        match &self.failure {
            Failure::InvalidDestination { kind, .. } => Some(*kind),
            Failure::Io { .. } | Failure::Committed { .. } => None,
        }
    }

    /// Return the publication phase represented by this error.
    #[must_use]
    pub const fn stage(&self) -> Stage {
        match &self.failure {
            Failure::InvalidDestination { stage, .. } => *stage,
            Failure::Io { stage, .. } => *stage,
            Failure::Committed { stage, .. } => *stage,
        }
    }

    /// Borrow the original operating-system error, when one caused this
    /// publication failure.
    #[must_use]
    pub fn io_error(&self) -> Option<&io::Error> {
        match &self.failure {
            Failure::Io { source, .. } | Failure::Committed { source, .. } => Some(source),
            Failure::InvalidDestination { .. } => None,
        }
    }

    /// Return the cached operating-system error kind, when one caused this
    /// publication failure.
    #[must_use]
    pub const fn io_error_kind(&self) -> Option<io::ErrorKind> {
        match &self.failure {
            Failure::Io { kind, .. } | Failure::Committed { kind, .. } => Some(*kind),
            Failure::InvalidDestination { .. } => None,
        }
    }

    /// Return whether replacement already committed before this error.
    #[must_use]
    pub const fn was_committed(&self) -> bool {
        matches!(&self.failure, Failure::Committed { .. })
    }

    /// Return whether replacement already committed before this error.
    #[must_use]
    pub const fn is_committed(&self) -> bool {
        self.was_committed()
    }
}

impl fmt::Debug for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.failure {
            Failure::InvalidDestination { kind, stage } => formatter
                .debug_struct("InvalidDestination")
                .field("kind", kind)
                .field("stage", stage)
                .finish(),
            Failure::Io { stage, kind, .. } => formatter
                .debug_struct("Io")
                .field("stage", stage)
                .field("kind", kind)
                .finish(),
            Failure::Committed { stage, kind, .. } => formatter
                .debug_struct("Committed")
                .field("stage", stage)
                .field("kind", kind)
                .finish(),
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.failure {
            Failure::InvalidDestination { kind, stage } => {
                write!(
                    formatter,
                    "archive publication destination is {kind} during {stage}"
                )
            },
            Failure::Io { stage, kind, .. } => write!(
                formatter,
                "archive publication failed during {stage} ({kind:?})"
            ),
            Failure::Committed { stage, kind, .. } => write!(
                formatter,
                "archive publication committed but post-replacement {stage} failed ({kind:?})"
            ),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match &self.failure {
            Failure::Io { source, .. } | Failure::Committed { source, .. } => Some(source),
            Failure::InvalidDestination { .. } => None,
        }
    }
}

/// Atomically replace `path` with bytes produced by `write`.
///
/// The callback receives a private, same-directory temporary file.  It is
/// called at most once.  Returning an error from the callback leaves the
/// destination untouched and makes a best-effort attempt to remove the
/// temporary file. Existing regular-file permissions represented by
/// [`std::fs::Permissions`] are copied to the replacement before it is
/// synchronized. On Unix this preserves ordinary mode bits but deliberately
/// clears set-user-ID and set-group-ID; ownership, ACLs, extended attributes,
/// security labels, and platform-specific metadata are outside this helper's
/// contract. Windows reparse-point destinations are rejected, and the
/// standard `Permissions`/read-only state is applied after replacement, but
/// no broader portable permission-preservation promise is made.
/// A newly created destination is forced to mode `0600` on Unix after the
/// callback returns. Existing broader modes are not applied until after
/// replacement through the returned open file handle, so the named staging
/// file remains private. If the callback introduces another buffering writer,
/// it must flush that writer and propagate any flush error before returning
/// `Ok`; this helper can flush only the supplied `File`.
/// The destination is checked again immediately before persistence so a
/// destination changed to a symbolic link or non-regular file is not followed
/// by the replacement operation.
///
/// The destination parent is resolved to one absolute path before the callback
/// runs, so a callback or another thread changing the process working directory
/// cannot redirect publication. Like other pathname-based atomic-save helpers,
/// this API assumes the resolved parent is controlled by the caller; it does
/// not defend against a same-identity adversary concurrently replacing names
/// inside a writable parent directory.
///
/// On Unix, the containing directory is synchronized after replacement when
/// the operating system supports that operation.  `InvalidInput` and
/// `Unsupported` directory-sync results are treated as an unsupported
/// durability capability. Permission application, file synchronization, or
/// parent synchronization that fails after replacement returns a committed
/// publication error (reported by [`Error::was_committed`]), because the new
/// destination bytes are already visible. On
/// Windows, `tempfile` supplies the platform atomic move/replacement and the
/// file synchronization plus rename sequence is the best-supported durability
/// available from this boundary; a strict directory-sync guarantee is not
/// claimed. Atomicity and durability ultimately depend on the destination
/// filesystem; no stronger guarantee is made for network or userspace
/// filesystems whose rename or synchronization semantics are weaker than the
/// host platform's local filesystem contract.
///
/// # Errors
///
/// Filesystem failures are converted to `E` through [`From<Error>`].  A
/// callback error is returned unchanged, so callback diagnostics remain
/// caller-owned.  Publication errors retain their operating-system source
/// behind accessors while keeping their default formatting content-free.
pub fn replace_with<E>(path: &Path, write: impl FnOnce(&mut File) -> Result<(), E>) -> Result<(), E>
where
    E: From<Error>,
{
    replace_with_impl(path, write, sync_parent)
}

fn replace_with_impl<E, S>(
    path: &Path,
    write: impl FnOnce(&mut File) -> Result<(), E>,
    sync: S,
) -> Result<(), E>
where
    E: From<Error>,
    S: FnOnce(&Path) -> io::Result<()>,
{
    ensure_publication_supported().map_err(E::from)?;
    let destination = resolve_destination(path).map_err(E::from)?;
    let path = destination.as_path();
    let parent = path.parent().ok_or_else(|| {
        E::from(Error::invalid_destination(
            DestinationKind::MissingFilename,
            Stage::ResolveParent,
        ))
    })?;
    let existing_permissions =
        inspect_destination(path, Stage::ValidateDestination).map_err(E::from)?;
    let mut replacement_permissions = existing_permissions.clone();

    // NamedTempFile defaults to mode 0600 on Unix and keeps the generated
    // pathname private to this operation.  Placing it beside the destination
    // also makes persist a same-filesystem rename.
    let mut temporary = tempfile::Builder::new()
        .prefix(".litchi-iwa-")
        .suffix(".tmp")
        .tempfile_in(parent)
        .map_err(|source| E::from(Error::io(Stage::CreateTemporary, source)))?;

    let staging_result = (|| {
        write(temporary.as_file_mut())?;

        // A callback receives the File because package writers stream into it,
        // but publication never inherits a broad mode onto the visible
        // staging pathname. Reset Unix staging permissions before any final
        // flush/sync so another principal cannot modify the inode pre-rename.
        if let Some(permissions) = private_staging_permissions() {
            temporary
                .as_file()
                .set_permissions(permissions)
                .map_err(|source| E::from(Error::io(Stage::SecureTemporary, source)))?;
        }

        // Flush the supplied File separately from sync_all. A callback that
        // adds another buffering layer must flush that layer before returning.
        temporary
            .as_file_mut()
            .flush()
            .map_err(|source| E::from(Error::io(Stage::Flush, source)))?;

        temporary
            .as_file_mut()
            .sync_all()
            .map_err(|source| E::from(Error::io(Stage::SyncFile, source)))?;

        // Re-read the destination after all callback output is complete. This
        // rejects a destination changed to a link or another file kind while
        // output was prepared. If a missing destination appeared, retain its
        // current permissions for post-replacement application.
        let current_permissions =
            inspect_destination(path, Stage::RevalidateDestination).map_err(E::from)?;
        if current_permissions.is_some() {
            replacement_permissions = current_permissions;
        }
        Ok(())
    })();
    if let Err(error) = staging_result {
        // Never unlink the generated name unless it still identifies the open
        // staging file. A callback may have replaced that directory entry.
        cleanup_temporary(temporary);
        return Err(error);
    }

    match temporary_path_matches_open_file(&temporary) {
        Ok(true) => {},
        Ok(false) => {
            // The generated pathname no longer identifies our open staging
            // file. Disarm tempfile cleanup so dropping this handle cannot
            // remove a replacement entry owned by somebody else.
            temporary.disable_cleanup(true);
            return Err(E::from(Error::io(
                Stage::RevalidateTemporary,
                io::Error::other("temporary publication pathname changed identity"),
            )));
        },
        Err(error) => {
            // Failure to authenticate the pathname is not permission to
            // unlink it: it may now identify an unrelated entry.
            temporary.disable_cleanup(true);
            return Err(E::from(error));
        },
    }
    let published = match temporary.persist(path) {
        Ok(file) => file,
        Err(failure) => {
            let source = Error::io(Stage::Persist, failure.error);
            cleanup_temporary(failure.file);
            return Err(E::from(source));
        },
    };

    // Apply any inherited destination mode only after the atomic replacement,
    // using the still-open handle returned by persist. The staging pathname
    // therefore never becomes group/other-readable or writable. Failures from
    // this point are committed failures because the new bytes are visible.
    if let Some(permissions) = publication_permissions(replacement_permissions.as_ref()) {
        published
            .set_permissions(permissions)
            .map_err(|source| E::from(Error::committed(Stage::PreservePermissions, source)))?;
        published
            .sync_all()
            .map_err(|source| E::from(Error::committed(Stage::SyncFile, source)))?;
    }

    if let Err(error) = sync(parent) {
        if matches!(
            error.kind(),
            io::ErrorKind::InvalidInput | io::ErrorKind::Unsupported
        ) {
            return Ok(());
        }
        return Err(E::from(Error::committed(Stage::SyncParent, error)));
    }
    Ok(())
}

fn cleanup_temporary(mut temporary: tempfile::NamedTempFile) {
    match temporary_path_matches_open_file(&temporary) {
        Ok(true) => {
            let _ = temporary.close();
        },
        Ok(false) | Err(_) => {
            // A missing or unauthenticated pathname may now belong to another
            // process. Leaking our unlinked handle is safer than deleting an
            // entry whose identity cannot be proved.
            temporary.disable_cleanup(true);
        },
    }
}

fn temporary_path_matches_open_file(temporary: &tempfile::NamedTempFile) -> Result<bool, Error> {
    let path_metadata = fs::symlink_metadata(temporary.path())
        .map_err(|source| Error::io(Stage::RevalidateTemporary, source))?;
    if path_metadata.file_type().is_symlink() || !path_metadata.is_file() {
        return Ok(false);
    }

    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;

        let open_metadata = temporary
            .as_file()
            .metadata()
            .map_err(|source| Error::io(Stage::RevalidateTemporary, source))?;
        Ok(
            open_metadata.dev() == path_metadata.dev()
                && open_metadata.ino() == path_metadata.ino(),
        )
    }
    #[cfg(windows)]
    {
        let open_file = temporary
            .as_file()
            .try_clone()
            .and_then(same_file::Handle::from_file)
            .map_err(|source| Error::io(Stage::RevalidateTemporary, source))?;
        let named_file = same_file::Handle::from_path(temporary.path())
            .map_err(|source| Error::io(Stage::RevalidateTemporary, source))?;
        Ok(open_file == named_file)
    }
    #[cfg(not(any(unix, windows)))]
    {
        Ok(false)
    }
}

fn ensure_publication_supported() -> Result<(), Error> {
    #[cfg(any(unix, windows))]
    {
        Ok(())
    }
    #[cfg(not(any(unix, windows)))]
    {
        Err(Error::io(
            Stage::ValidateDestination,
            io::Error::from(io::ErrorKind::Unsupported),
        ))
    }
}

fn resolve_destination(path: &Path) -> Result<PathBuf, Error> {
    let filename = path.file_name().ok_or_else(|| {
        Error::invalid_destination(DestinationKind::MissingFilename, Stage::ValidateDestination)
    })?;
    let supplied_parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let resolved_parent = fs::canonicalize(supplied_parent)
        .map_err(|source| Error::io(Stage::ResolveParent, source))?;
    Ok(resolved_parent.join(filename))
}

fn inspect_destination(path: &Path, stage: Stage) -> Result<Option<Permissions>, Error> {
    match fs::symlink_metadata(path) {
        Ok(metadata) if metadata.file_type().is_symlink() => {
            Err(Error::invalid_destination(DestinationKind::Symlink, stage))
        },
        Ok(metadata) if metadata_is_reparse_point(&metadata) => Err(Error::invalid_destination(
            DestinationKind::ReparsePoint,
            stage,
        )),
        Ok(metadata) if !metadata.is_file() => Err(Error::invalid_destination(
            DestinationKind::NonRegular,
            stage,
        )),
        Ok(metadata) => Ok(Some(metadata.permissions())),
        Err(error) if error.kind() == io::ErrorKind::NotFound => Ok(None),
        Err(source) => Err(Error::io(stage, source)),
    }
}

#[cfg(unix)]
fn private_staging_permissions() -> Option<Permissions> {
    use std::os::unix::fs::PermissionsExt;

    Some(Permissions::from_mode(0o600))
}

#[cfg(not(unix))]
fn private_staging_permissions() -> Option<Permissions> {
    None
}

#[cfg(unix)]
fn publication_permissions(existing: Option<&Permissions>) -> Option<Permissions> {
    use std::os::unix::fs::PermissionsExt;

    let mode = existing.map_or(0o600, PermissionsExt::mode) & !0o6000;
    Some(Permissions::from_mode(mode))
}

#[cfg(not(unix))]
fn publication_permissions(existing: Option<&Permissions>) -> Option<Permissions> {
    existing.cloned()
}

fn metadata_is_reparse_point(metadata: &fs::Metadata) -> bool {
    #[cfg(windows)]
    {
        use std::os::windows::fs::MetadataExt;

        const FILE_ATTRIBUTE_REPARSE_POINT: u32 = 0x0400;
        metadata.file_attributes() & FILE_ATTRIBUTE_REPARSE_POINT != 0
    }
    #[cfg(not(windows))]
    {
        let _ = metadata;
        false
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
            // Some Unix filesystems expose no directory-sync operation.  The
            // replacement remains committed; the caller treats this as an
            // unsupported capability rather than a publication failure.
            Err(error)
        },
        Err(error) => Err(error),
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
        reason = "test assertions panic on failure by design"
    )]

    use super::*;

    fn assert_send_sync_debug<T: Send + Sync + fmt::Debug>() {}

    #[derive(Debug)]
    enum TypedError {
        Publication(Error),
        Callback,
    }

    impl From<Error> for TypedError {
        fn from(error: Error) -> Self {
            Self::Publication(error)
        }
    }

    #[test]
    fn typed_publication_values_are_send_sync_and_redacted() {
        assert_send_sync_debug::<DestinationKind>();
        assert_send_sync_debug::<Stage>();
        assert_send_sync_debug::<Error>();

        let error =
            Error::invalid_destination(DestinationKind::Symlink, Stage::ValidateDestination);
        let debug = format!("{error:?}");
        let display = error.to_string();
        assert!(!debug.contains("secret/path"));
        assert!(!display.contains("secret/path"));
        assert_eq!(error.destination_kind(), Some(DestinationKind::Symlink));
        assert_eq!(error.stage(), Stage::ValidateDestination);
        assert!(!error.is_committed());
    }

    #[test]
    fn missing_filename_is_rejected_without_running_callback() {
        let mut called = false;
        let result = replace_with::<TypedError>(Path::new(""), |_file| {
            called = true;
            Ok(())
        });

        assert!(!called);
        match result {
            Err(TypedError::Publication(error)) => {
                assert_eq!(
                    error.destination_kind(),
                    Some(DestinationKind::MissingFilename)
                );
            },
            _ => panic!("missing filename should be rejected"),
        }
    }

    #[test]
    fn relative_destination_is_fixed_to_an_absolute_parent() {
        let destination =
            resolve_destination(Path::new("document.iwa")).expect("resolve relative destination");
        let expected_parent = fs::canonicalize(".").expect("resolve current directory");

        assert!(destination.is_absolute());
        assert_eq!(destination.parent(), Some(expected_parent.as_path()));
        assert_eq!(
            destination.file_name(),
            Some(std::ffi::OsStr::new("document.iwa"))
        );
    }

    #[test]
    fn callback_failure_preserves_destination_and_removes_tempfile() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.iwa");
        fs::write(&destination, b"original").expect("seed destination");

        let result = replace_with::<TypedError>(&destination, |file| {
            file.write_all(b"partial").expect("write temporary");
            Err(TypedError::Callback)
        });

        assert!(matches!(result, Err(TypedError::Callback)));
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
    fn callback_failure_does_not_delete_a_replaced_temporary_path() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.iwa");
        fs::write(&destination, b"original").expect("seed destination");
        let mut replacement_path = None;

        let result = replace_with::<TypedError>(&destination, |file| {
            file.write_all(b"partial").expect("write staging file");
            let temporary_path = fs::read_dir(directory.path())
                .expect("list destination parent")
                .map(|entry| entry.expect("temporary entry").path())
                .find(|path| {
                    path.file_name().is_some_and(|name| {
                        let name = name.to_string_lossy();
                        name.starts_with(".litchi-iwa-") && name.ends_with(".tmp")
                    })
                })
                .expect("private staging pathname");
            fs::remove_file(&temporary_path).expect("unlink staging pathname");
            fs::write(&temporary_path, b"substitute").expect("replace staging pathname");
            replacement_path = Some(temporary_path);
            Err(TypedError::Callback)
        });

        assert!(matches!(result, Err(TypedError::Callback)));
        assert_eq!(
            fs::read(&destination).expect("read destination"),
            b"original"
        );
        let replacement_path = replacement_path.expect("replacement pathname");
        assert_eq!(
            fs::read(replacement_path).expect("read substitute"),
            b"substitute"
        );
    }

    #[test]
    fn replaced_temporary_path_is_not_published_or_deleted() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.iwa");
        fs::write(&destination, b"original").expect("seed destination");
        let mut replacement_path = None;

        let result = replace_with::<TypedError>(&destination, |file| {
            file.write_all(b"authored").expect("write staging file");
            let temporary_path = fs::read_dir(directory.path())
                .expect("list destination parent")
                .map(|entry| entry.expect("temporary entry").path())
                .find(|path| {
                    path.file_name().is_some_and(|name| {
                        let name = name.to_string_lossy();
                        name.starts_with(".litchi-iwa-") && name.ends_with(".tmp")
                    })
                })
                .expect("private staging pathname");
            fs::remove_file(&temporary_path).expect("unlink staging pathname");
            fs::write(&temporary_path, b"substitute").expect("replace staging pathname");
            replacement_path = Some(temporary_path);
            Ok(())
        });

        let error = match result {
            Err(TypedError::Publication(error)) => error,
            _ => panic!("replaced staging pathname should be rejected"),
        };
        assert_eq!(error.stage(), Stage::RevalidateTemporary);
        assert_eq!(
            fs::read(&destination).expect("read destination"),
            b"original"
        );
        let replacement_path = replacement_path.expect("replacement pathname");
        assert_eq!(
            fs::read(replacement_path).expect("read substitute"),
            b"substitute"
        );
    }

    #[test]
    fn successful_publication_replaces_destination() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.iwa");
        fs::write(&destination, b"old").expect("seed destination");

        replace_with::<TypedError>(&destination, |file| {
            file.write_all(b"new").expect("write temporary");
            Ok(())
        })
        .expect("publication");

        assert_eq!(fs::read(destination).expect("read destination"), b"new");
    }

    #[test]
    fn missing_destination_is_created() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.iwa");

        replace_with::<TypedError>(&destination, |file| {
            file.write_all(b"new").expect("write temporary");
            Ok(())
        })
        .expect("publication");

        assert_eq!(fs::read(destination).expect("read destination"), b"new");
    }

    #[cfg(unix)]
    #[test]
    fn symbolic_link_destination_is_rejected_without_touching_target() {
        use std::os::unix::fs::symlink;

        let directory = tempfile::tempdir().expect("temporary directory");
        let target = directory.path().join("target.iwa");
        let destination = directory.path().join("document.iwa");
        fs::write(&target, b"target").expect("seed target");
        symlink(&target, &destination).expect("create symlink");

        let result = replace_with::<TypedError>(&destination, |_file| Ok(()));

        match result {
            Err(TypedError::Publication(error)) => {
                assert_eq!(error.destination_kind(), Some(DestinationKind::Symlink));
            },
            _ => panic!("symbolic link should be rejected"),
        }
        assert_eq!(fs::read(&target).expect("read target"), b"target");
        assert!(
            fs::symlink_metadata(&destination)
                .expect("destination metadata")
                .file_type()
                .is_symlink()
        );
    }

    #[cfg(unix)]
    #[test]
    fn destination_changed_to_a_symbolic_link_is_rejected_before_persist() {
        use std::os::unix::fs::symlink;

        let directory = tempfile::tempdir().expect("temporary directory");
        let target = directory.path().join("target.iwa");
        let destination = directory.path().join("document.iwa");
        fs::write(&target, b"target").expect("seed target");
        fs::write(&destination, b"original").expect("seed destination");

        let result = replace_with::<TypedError>(&destination, |file| {
            file.write_all(b"new").expect("write temporary");
            fs::remove_file(&destination).expect("remove original destination");
            symlink(&target, &destination).expect("replace destination with symlink");
            Ok(())
        });

        let error = match result {
            Err(TypedError::Publication(error)) => error,
            _ => panic!("raced symbolic link should be rejected"),
        };
        assert_eq!(error.destination_kind(), Some(DestinationKind::Symlink));
        assert_eq!(error.stage(), Stage::RevalidateDestination);
        assert_eq!(fs::read(&target).expect("read target"), b"target");
        assert!(
            fs::symlink_metadata(&destination)
                .expect("destination metadata")
                .file_type()
                .is_symlink()
        );
        assert_eq!(
            fs::read_dir(directory.path())
                .expect("list temporary directory")
                .count(),
            2
        );
    }

    #[test]
    fn nonregular_destination_is_rejected() {
        let directory = tempfile::tempdir().expect("temporary directory");

        let result = replace_with::<TypedError>(directory.path(), |_file| Ok(()));

        match result {
            Err(TypedError::Publication(error)) => {
                assert_eq!(error.destination_kind(), Some(DestinationKind::NonRegular));
            },
            _ => panic!("non-regular destination should be rejected"),
        }
    }

    #[cfg(unix)]
    #[test]
    fn existing_permissions_are_preserved() {
        use std::os::unix::fs::PermissionsExt;

        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.iwa");
        fs::write(&destination, b"old").expect("seed destination");
        fs::set_permissions(&destination, Permissions::from_mode(0o640))
            .expect("set destination permissions");

        replace_with::<TypedError>(&destination, |file| {
            file.write_all(b"new").expect("write temporary");
            Ok(())
        })
        .expect("publication");

        let mode = fs::metadata(destination)
            .expect("destination metadata")
            .permissions()
            .mode()
            & 0o777;
        assert_eq!(mode, 0o640);
    }

    #[cfg(unix)]
    #[test]
    fn preserved_permissions_clear_privileged_execution_bits() {
        use std::os::unix::fs::PermissionsExt;

        let permissions = publication_permissions(Some(&Permissions::from_mode(0o6755)))
            .expect("Unix publication permissions");

        assert_eq!(permissions.mode() & 0o7777, 0o755);
    }

    #[cfg(unix)]
    #[test]
    fn staging_permissions_do_not_inherit_a_shared_destination_mode() {
        use std::os::unix::fs::PermissionsExt;

        let staging = private_staging_permissions().expect("Unix staging permissions");
        let published = publication_permissions(Some(&Permissions::from_mode(0o666)))
            .expect("Unix publication permissions");

        assert_eq!(staging.mode() & 0o777, 0o600);
        assert_eq!(published.mode() & 0o777, 0o666);
    }

    #[cfg(unix)]
    #[test]
    fn new_destination_is_forced_to_private_permissions() {
        use std::os::unix::fs::PermissionsExt;

        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.iwa");

        replace_with::<TypedError>(&destination, |file| {
            file.write_all(b"new").expect("write temporary");
            file.set_permissions(Permissions::from_mode(0o777))
                .expect("broaden staging permissions");
            Ok(())
        })
        .expect("publication");

        let mode = fs::metadata(destination)
            .expect("destination metadata")
            .permissions()
            .mode()
            & 0o777;
        assert_eq!(mode, 0o600);
    }

    #[cfg(unix)]
    #[test]
    fn unreadable_existing_mode_is_preserved_without_blocking_publication() {
        use std::os::unix::fs::PermissionsExt;

        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.iwa");
        fs::write(&destination, b"old").expect("seed destination");
        fs::set_permissions(&destination, Permissions::from_mode(0o000))
            .expect("set unreadable destination permissions");

        replace_with::<TypedError>(&destination, |file| {
            file.write_all(b"new").expect("write temporary");
            Ok(())
        })
        .expect("publication");

        let mode = fs::metadata(&destination)
            .expect("destination metadata")
            .permissions()
            .mode()
            & 0o777;
        assert_eq!(mode, 0o000);
        fs::set_permissions(&destination, Permissions::from_mode(0o600))
            .expect("restore readable destination permissions");
        assert_eq!(fs::read(destination).expect("read destination"), b"new");
    }

    #[cfg(unix)]
    #[test]
    fn parent_sync_failure_reports_committed_after_replacement() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.iwa");
        fs::write(&destination, b"old").expect("seed destination");

        let result = replace_with_impl::<TypedError, _>(
            &destination,
            |file| {
                file.write_all(b"new").expect("write temporary");
                Ok(())
            },
            |_parent| {
                Err(io::Error::new(
                    io::ErrorKind::PermissionDenied,
                    "private parent-sync detail",
                ))
            },
        );

        let error = match result {
            Err(TypedError::Publication(error)) => error,
            _ => panic!("parent sync should fail after replacement"),
        };
        assert_eq!(fs::read(destination).expect("read destination"), b"new");
        let debug = format!("{error:?}");
        assert!(!debug.contains("private parent-sync detail"));
        assert!(error.was_committed());
        assert_eq!(error.stage(), Stage::SyncParent);
        assert_eq!(error.io_error_kind(), Some(io::ErrorKind::PermissionDenied));
        assert_eq!(
            error
                .io_error()
                .expect("retained parent-sync error")
                .to_string(),
            "private parent-sync detail"
        );
    }

    #[test]
    fn unsupported_parent_sync_is_accepted() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("document.iwa");

        replace_with_impl::<TypedError, _>(
            &destination,
            |file| {
                file.write_all(b"new").expect("write temporary");
                Ok(())
            },
            |_parent| Err(io::Error::from(io::ErrorKind::Unsupported)),
        )
        .expect("unsupported parent sync does not undo replacement");

        assert_eq!(fs::read(destination).expect("read destination"), b"new");
    }
}
