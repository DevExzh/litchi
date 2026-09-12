//! Shared failure-atomic filesystem publication for flat ODT snapshots.
//!
//! Unix uses same-directory rename followed by parent-directory
//! synchronization. Windows uses `tempfile`'s same-directory replacement
//! primitive and file synchronization; parent-directory durability is not
//! claimed there. Other targets fail before creating a staging file.

use std::io::{self, Write};
use std::path::Path;
use std::sync::Arc;

#[cfg(windows)]
use same_file::Handle;
use tempfile::NamedTempFile;

#[cfg(unix)]
use std::os::unix::fs::MetadataExt;

use litchi_core::{Error, ExecutionContext, ExecutionError, Resource, ResourceLimit, Result};

/// Write `bytes` through a sibling temporary and publish them atomically,
/// checking the caller's execution context throughout staging.
///
/// The destination is checked before any temporary is created. The temporary
/// is flushed, synchronized, length-checked, and identity-checked before
/// publication. A pre-publication failure leaves the existing destination
/// untouched when the temporary pathname can be authenticated. Once the
/// destination has been replaced, a parent-sync failure returns the public
/// [`Error::Committed`] variant; callers must not retry it as an ordinary
/// pre-publication failure.
pub(crate) fn save_with_context(
    path: &Path,
    bytes: &[u8],
    max_output_bytes: usize,
    output_scope: &'static str,
    destination_scope: &'static str,
    context: &ExecutionContext,
) -> Result<()> {
    save_with_sync_context(
        path,
        bytes,
        max_output_bytes,
        output_scope,
        destination_scope,
        Some(context),
        sync_parent,
    )
}

#[cfg(all(test, unix))]
fn save_with_sync(
    path: &Path,
    bytes: &[u8],
    max_output_bytes: usize,
    output_scope: &'static str,
    destination_scope: &'static str,
    sync: impl FnOnce(&Path) -> io::Result<()>,
) -> Result<()> {
    save_with_sync_context(
        path,
        bytes,
        max_output_bytes,
        output_scope,
        destination_scope,
        None,
        sync,
    )
}

fn save_with_sync_context(
    path: &Path,
    bytes: &[u8],
    max_output_bytes: usize,
    output_scope: &'static str,
    destination_scope: &'static str,
    context: Option<&ExecutionContext>,
    sync: impl FnOnce(&Path) -> io::Result<()>,
) -> Result<()> {
    check_context(context)?;
    consume_work(context)?;
    let parent = path
        .parent()
        .filter(|value| !value.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    validate_destination(path, destination_scope)?;

    if bytes.len() > max_output_bytes {
        return Err(resource_limit_error(
            Resource::OutputBytes,
            bytes.len(),
            max_output_bytes,
            output_scope,
        ));
    }

    ensure_publication_supported()?;

    let mut temporary = Some(create_owned_sibling_temp(parent)?);
    let mut published = false;
    let write_result = (|| -> Result<()> {
        // Keep every post-creation failure inside the guarded closure so the
        // identity-aware cleanup path owns the temporary pathname.
        check_context(context)?;
        consume_work(context)?;
        let Some(staging) = temporary.as_mut() else {
            return Err(Error::Other(
                "flat ODT temporary ownership was lost before publication".to_string(),
            ));
        };
        // Keep cancellation responsive for large snapshots and consume work
        // in the same bounded chunks used for staging. The final check below
        // is the publication fence; no cancellation check occurs after the
        // replacement has committed.
        const WRITE_CHUNK_BYTES: usize = 64 * 1024;
        let mut offset = 0;
        while offset < bytes.len() {
            check_context(context)?;
            consume_work(context)?;
            let end = offset.saturating_add(WRITE_CHUNK_BYTES).min(bytes.len());
            staging.as_file_mut().write_all(&bytes[offset..end])?;
            offset = end;
        }
        check_context(context)?;
        staging.as_file_mut().flush()?;
        check_context(context)?;
        staging.as_file_mut().sync_all()?;
        check_context(context)?;
        let expected_len = u64::try_from(bytes.len()).map_err(|_error| {
            resource_limit_error(
                Resource::OutputBytes,
                bytes.len(),
                max_output_bytes,
                output_scope,
            )
        })?;
        let actual_len = staging.as_file().metadata()?.len();
        if actual_len != expected_len {
            return Err(ResourceLimit {
                resource: Resource::OutputBytes,
                observed: actual_len,
                limit: expected_len,
                scope: Arc::from(output_scope),
            }
            .into());
        }

        if !temporary_path_matches_open_file(staging)? {
            return Err(Error::Other(
                "flat ODT temporary publication pathname changed identity".to_string(),
            ));
        }
        validate_destination(path, destination_scope)?;
        // This is the last cancellable point before replacement. Once
        // `publish_staging` returns, the destination may contain the new
        // snapshot and a later durability error is reported as committed.
        check_context(context)?;
        consume_work(context)?;

        let Some(staging) = temporary.take() else {
            return Err(Error::Other(
                "flat ODT temporary ownership was lost before publication".to_string(),
            ));
        };
        let published_file = publish_staging(staging, path)?;
        // `persist` has completed the replacement at this point. Keep this
        // state before any later durability work so cleanup cannot unlink a
        // pathname that now belongs to the published destination.
        published = true;
        drop(published_file);
        sync(parent).map_err(committed_sync_error)?;
        Ok(())
    })();
    if write_result.is_err() && !published {
        if let Some(staging) = temporary.take() {
            cleanup_temporary(staging);
        }
    }
    write_result
}

fn check_context(context: Option<&ExecutionContext>) -> Result<()> {
    let Some(context) = context else {
        return Ok(());
    };
    context.check().map_err(map_execution_error)
}

fn consume_work(context: Option<&ExecutionContext>) -> Result<()> {
    let Some(context) = context else {
        return Ok(());
    };
    context
        .consume(Resource::Work, 1)
        .map_err(map_execution_error)
}

fn map_execution_error(error: ExecutionError) -> Error {
    match error {
        ExecutionError::ResourceLimit(limit) => Error::ResourceLimit(limit),
        ExecutionError::Cancelled => Error::Other("flat ODT publication cancelled".to_string()),
        other => Error::Other(format!("flat ODT publication context failed: {other}")),
    }
}

fn ensure_publication_supported() -> Result<()> {
    #[cfg(any(unix, windows))]
    {
        Ok(())
    }
    #[cfg(not(any(unix, windows)))]
    {
        Err(Error::Unsupported(
            "atomic flat ODT publication is unavailable on this platform".to_string(),
        ))
    }
}

fn validate_destination(path: &Path, scope: &'static str) -> Result<()> {
    match std::fs::symlink_metadata(path) {
        Ok(metadata) if metadata.file_type().is_symlink() => {
            Err(resource_limit_error(Resource::Objects, 1, 0, scope))
        },
        Ok(metadata) if !metadata.is_file() => {
            Err(resource_limit_error(Resource::Objects, 1, 0, scope))
        },
        Ok(_) => Ok(()),
        Err(error) if error.kind() == io::ErrorKind::NotFound => Ok(()),
        Err(error) => Err(error.into()),
    }
}

fn create_owned_sibling_temp(parent: &Path) -> Result<NamedTempFile> {
    tempfile::Builder::new()
        .prefix(".litchi-")
        .suffix(".tmp")
        .tempfile_in(parent)
        .map_err(Into::into)
}

fn publish_staging(staging: NamedTempFile, destination: &Path) -> io::Result<std::fs::File> {
    match staging.persist(destination) {
        Ok(file) => Ok(file),
        Err(failure) => {
            let source = failure.error;
            cleanup_temporary(failure.file);
            Err(source)
        },
    }
}

fn temporary_path_matches_open_file(temporary: &NamedTempFile) -> Result<bool> {
    let path_metadata = match std::fs::symlink_metadata(temporary.path()) {
        Ok(metadata) => metadata,
        Err(error) if error.kind() == io::ErrorKind::NotFound => return Ok(false),
        Err(error) => return Err(error.into()),
    };
    if path_metadata.file_type().is_symlink() || !path_metadata.is_file() {
        return Ok(false);
    }

    #[cfg(unix)]
    {
        let open_metadata = temporary.as_file().metadata()?;
        return Ok(open_metadata.dev() == path_metadata.dev()
            && open_metadata.ino() == path_metadata.ino());
    }
    #[cfg(windows)]
    {
        let open_file = temporary.as_file().try_clone()?;
        let open_handle = Handle::from_file(open_file)?;
        let named_handle = Handle::from_path(temporary.path())?;
        return Ok(open_handle == named_handle);
    }
    #[cfg(not(any(unix, windows)))]
    {
        Ok(false)
    }
}

fn cleanup_temporary(mut temporary: NamedTempFile) {
    match temporary_path_matches_open_file(&temporary) {
        Ok(true) => {
            drop(temporary.close());
        },
        Ok(false) | Err(_) => {
            // A missing or unauthenticated pathname may now belong to another
            // process. Leaking our open handle is safer than deleting an entry
            // whose identity cannot be proved.
            temporary.disable_cleanup(true);
        },
    }
}

#[cfg(unix)]
fn sync_parent(parent: &Path) -> io::Result<()> {
    std::fs::File::open(parent)?.sync_all()
}

#[cfg(not(unix))]
fn sync_parent(_parent: &Path) -> io::Result<()> {
    Ok(())
}

fn committed_sync_error(source: io::Error) -> Error {
    Error::Committed(source)
}

fn resource_limit_error(
    resource: Resource,
    observed: usize,
    limit: usize,
    scope: &'static str,
) -> Error {
    ResourceLimit {
        resource,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        limit: u64::try_from(limit).unwrap_or(u64::MAX),
        scope: Arc::from(scope),
    }
    .into()
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions panic on failure by design"
    )]

    use super::*;
    use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits as BudgetLimits};
    use std::num::{NonZeroU64, NonZeroUsize};

    fn test_context(work: u64) -> (Budget, CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            "flat atomic publication test",
            BudgetLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work),
        );
        let (source, token) = CancellationSource::pair();
        let limits = ExecutionLimits::new(
            NonZeroUsize::new(1).expect("one worker"),
            NonZeroUsize::new(1).expect("one task"),
            NonZeroU64::new(u64::MAX).expect("nonzero bytes"),
            0,
        )
        .expect("valid execution limits");
        (
            budget.clone(),
            source,
            ExecutionContext::new(budget, token, limits),
        )
    }

    #[test]
    fn committed_sync_error_is_programmatically_distinguishable() {
        let source = io::Error::new(io::ErrorKind::PermissionDenied, "injected sync failure");
        let error = committed_sync_error(source);
        assert!(matches!(error, Error::Committed(_)));
        assert!(error.to_string().contains("publication committed"));
    }

    #[test]
    fn cancelled_publication_leaves_destination_and_staging_untouched() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("template.fott");
        std::fs::write(&destination, b"old").expect("seed destination");
        let (_budget, cancellation, context) = test_context(u64::MAX);
        cancellation.cancel();

        let result = save_with_context(
            &destination,
            b"new",
            3,
            "flat ODT test output",
            "flat ODT test destination",
            &context,
        );

        assert!(matches!(result, Err(Error::Other(message)) if message.contains("cancelled")));
        assert_eq!(
            std::fs::read(&destination).expect("destination remains"),
            b"old"
        );
        assert_eq!(
            std::fs::read_dir(directory.path())
                .expect("directory remains readable")
                .count(),
            1,
            "cancelled publication must not leave a staging pathname"
        );
    }

    #[test]
    fn post_creation_context_failure_uses_identity_aware_cleanup() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("template.fott");
        std::fs::write(&destination, b"old").expect("seed destination");
        let (_budget, _cancellation, context) = test_context(1);

        let result = save_with_context(
            &destination,
            b"new",
            3,
            "flat ODT test output",
            "flat ODT test destination",
            &context,
        );

        assert!(
            matches!(result, Err(Error::ResourceLimit(limit)) if limit.resource == Resource::Work)
        );
        assert_eq!(
            std::fs::read(&destination).expect("destination remains"),
            b"old"
        );
        assert_eq!(
            std::fs::read_dir(directory.path())
                .expect("directory remains readable")
                .count(),
            1,
            "post-creation failure must clean only the owned staging pathname"
        );
    }

    #[cfg(unix)]
    #[test]
    fn parent_sync_failure_reports_committed_after_replacement() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("template.fott");
        std::fs::write(&destination, b"old").expect("seed destination");

        let result = save_with_sync(
            &destination,
            b"new",
            3,
            "flat ODT test output",
            "flat ODT test destination",
            |_parent| {
                Err(io::Error::new(
                    io::ErrorKind::PermissionDenied,
                    "injected directory sync failure",
                ))
            },
        );

        assert!(matches!(
            result,
            Err(Error::Committed(source))
                if source.kind() == io::ErrorKind::PermissionDenied
        ));
        assert_eq!(
            std::fs::read(&destination).expect("published destination"),
            b"new"
        );
    }

    #[cfg(unix)]
    #[test]
    fn cleanup_never_unlinks_a_replaced_temporary_path() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let temporary = tempfile::Builder::new()
            .prefix(".litchi-test-")
            .tempfile_in(directory.path())
            .expect("staging file");
        let path = temporary.path().to_owned();
        std::fs::remove_file(&path).expect("remove staging pathname");
        std::fs::write(&path, b"replacement").expect("replacement pathname");

        cleanup_temporary(temporary);

        assert_eq!(
            std::fs::read(path).expect("replacement remains"),
            b"replacement"
        );
    }
}
