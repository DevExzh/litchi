//! App-authored iWork package directories (the bundle form), shared by the
//! Pages, Keynote and Numbers owners (`IWORK_API_REQUESTS.md` request 16,
//! records 0299, 0301 and 0306).
//!
//! An iWork application can save a document as a directory bundle:
//! `Index.zip` holds the IWA components (already under `Index/`), and the
//! metadata sidecars, previews and `Data/` media sit beside it as plain
//! files. The single-file form is one ZIP holding the same members.
//!
//! - [`capture`] reads a bundle once, under caller-selected
//!   [`CaptureLimits`], and converts it: the `Index.zip` members first (their
//!   records copied as stored), then every other file of the bundle as a
//!   stored member, in byte order of its relative name, so the conversion is
//!   deterministic. The walk refuses symbolic links and special files, bounds
//!   the depth, the member count and every byte figure, and re-checks each
//!   file's length after reading it, so a bundle that changes while it is
//!   captured is refused rather than half-read.
//! - [`save`] writes a package back as a bundle through the directory
//!   publication boundary ([`crate::publication::replace_directory_with`]):
//!   `Index.zip` holds the `Index/` members with their records unchanged, and
//!   every other member is a plain file. Capturing the saved bundle gives back
//!   the package's exact single-file bytes.

use std::fs;
use std::io::{self, Read};
use std::path::Path;

use crate::package::{Catalog, EntryInsertion};
use crate::{Error, Limits};

/// The deepest directory nesting a bundle may use below its root.
const MAX_DEPTH: usize = 16;
/// The IWA component archive at the root of a bundle.
pub const INDEX_ZIP: &str = "Index.zip";

/// The physical bounds of one bundle capture.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CaptureLimits {
    max_members: usize,
    max_entry_bytes: u64,
    max_input_bytes: u64,
}

impl CaptureLimits {
    /// Bounds on the number of files, the size of each file and the sum of
    /// all file sizes.
    #[must_use]
    pub const fn new(max_members: usize, max_entry_bytes: u64, max_input_bytes: u64) -> Self {
        Self {
            max_members,
            max_entry_bytes,
            max_input_bytes,
        }
    }
}

/// The capture dimension a bundle exceeded.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CaptureLimitKind {
    /// The number of files.
    Members,
    /// The size of one file.
    EntryBytes,
    /// The sum of all file sizes.
    InputBytes,
}

/// A failure to capture a package directory.
#[derive(Debug)]
#[non_exhaustive]
pub enum CaptureError {
    /// The directory is not a supported bundle; the message names why.
    Invalid(&'static str),
    /// A capture bound was exceeded.
    Limit {
        /// The exceeded dimension.
        kind: CaptureLimitKind,
        /// The observed figure.
        observed: u64,
        /// The configured maximum.
        maximum: u64,
    },
    /// A file's buffer could not be reserved.
    Allocation {
        /// The requested byte count.
        amount: usize,
    },
    /// A filesystem read failed.
    Io(io::Error),
    /// `Index.zip` is malformed, or the conversion's reassembly failed.
    Archive(Error),
}

impl std::fmt::Display for CaptureError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Invalid(message) => formatter.write_str(message),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "the package directory exceeds its {kind:?} limit ({observed} > {maximum})"
            ),
            Self::Allocation { amount } => {
                write!(
                    formatter,
                    "could not reserve {amount} bytes for a bundle file"
                )
            },
            Self::Io(_) => formatter.write_str("could not read the package directory"),
            Self::Archive(_) => formatter.write_str("could not convert the package directory"),
        }
    }
}

impl std::error::Error for CaptureError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(error) => Some(error),
            Self::Archive(error) => Some(error),
            Self::Invalid(_) | Self::Limit { .. } | Self::Allocation { .. } => None,
        }
    }
}

impl From<io::Error> for CaptureError {
    fn from(error: io::Error) -> Self {
        Self::Io(error)
    }
}

impl From<Error> for CaptureError {
    fn from(error: Error) -> Self {
        Self::Archive(error)
    }
}

/// The archive error of a capture failure, for owners that report physical
/// failures as [`Error`]: an invalid bundle is [`Error::InvalidBundle`], a
/// capture bound is the matching [`Error::Limit`].
impl From<CaptureError> for Error {
    fn from(error: CaptureError) -> Self {
        match error {
            CaptureError::Invalid(message) => Self::InvalidBundle(message.to_owned()),
            CaptureError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    CaptureLimitKind::Members => crate::LimitKind::Entries,
                    CaptureLimitKind::EntryBytes => crate::LimitKind::EntryBytes,
                    CaptureLimitKind::InputBytes => crate::LimitKind::InputBytes,
                },
                observed,
                maximum,
            },
            CaptureError::Allocation { amount } => Self::Allocation {
                resource: "package directory file",
                amount,
            },
            CaptureError::Io(error) => Self::Io(error),
            CaptureError::Archive(error) => error,
        }
    }
}

/// The capture bounds an archive profile implies: its member count, its
/// per-member size, and the smaller of its input and aggregate ceilings.
impl From<Limits> for CaptureLimits {
    fn from(limits: Limits) -> Self {
        Self::new(
            limits.max_entries(),
            limits.max_entry_bytes(),
            limits.max_input_bytes().min(limits.max_total_bytes()),
        )
    }
}

/// One captured bundle file: its `/`-separated relative name and bytes.
struct Member {
    name: String,
    data: Vec<u8>,
}

/// Whether `path` is a directory (not following a final symbolic link).
#[must_use]
pub fn is_directory(path: &Path) -> bool {
    fs::symlink_metadata(path).is_ok_and(|metadata| metadata.is_dir())
}

/// Capture the bundle at `root` and return its single-file package bytes.
/// `archive` is the physical profile `Index.zip` is read under; the
/// conversion's own output bound is derived from the captured length.
///
/// A bundle normally holds `Index.zip`. A bundle whose components sit loose
/// in an `Index/` folder instead (an older app layout) converts too: every
/// file, the `Index/` components included, becomes a stored member in byte
/// order of its relative name. Saving such a package with [`save`] writes
/// the `Index.zip` layout.
///
/// # Errors
///
/// [`CaptureError::Invalid`] for a bundle with neither `Index.zip` nor an
/// `Index/` folder, with both, with a symbolic link, a special file, a
/// non-UTF-8 name or nesting deeper than 16 levels; [`CaptureError::Limit`]
/// when a figure exceeds `limits`; [`CaptureError::Io`] for a read failure;
/// [`CaptureError::Archive`] when `Index.zip` is malformed or the
/// conversion fails.
pub fn capture(
    root: &Path,
    limits: CaptureLimits,
    archive: Limits,
) -> Result<Vec<u8>, CaptureError> {
    let mut members = Vec::new();
    let mut total = 0_u64;
    walk(root, "", 0, limits, &mut members, &mut total)?;
    members.sort_by(|left, right| left.name.as_bytes().cmp(right.name.as_bytes()));
    let loose = is_directory(&root.join("Index"));
    let index = members.iter().position(|member| member.name == INDEX_ZIP);
    let index = match (index, loose) {
        (Some(_), true) => {
            return Err(CaptureError::Invalid(
                "an iWork package directory holds both Index.zip and an Index folder",
            ));
        },
        (None, false) => {
            return Err(CaptureError::Invalid(
                "an iWork package directory has no Index.zip",
            ));
        },
        (None, true) => {
            let output =
                usize::try_from(total)
                    .unwrap_or(usize::MAX)
                    .saturating_add(members.iter().fold(64_usize, |bound, member| {
                        bound
                            .saturating_add(100)
                            .saturating_add(member.name.len().saturating_mul(2))
                    }));
            let archive = converted_limits(archive, output)?;
            let entries: Vec<(&str, &[u8])> = members
                .iter()
                .map(|member| (member.name.as_str(), member.data.as_slice()))
                .collect();
            return Ok(crate::package::to_bytes(entries, archive)?);
        },
        (Some(index), false) => index,
    };
    let index = members.remove(index);
    let catalog = Catalog::from_bytes(&index.data, archive)?;
    // The converted package is the captured bytes plus ZIP framing: a local
    // and a central record per inserted member (each under 100 bytes plus
    // the name twice) and one end record.
    let framing = members.iter().fold(64_usize, |bound, member| {
        bound
            .saturating_add(100)
            .saturating_add(member.name.len().saturating_mul(2))
    });
    let output = usize::try_from(total)
        .unwrap_or(usize::MAX)
        .saturating_add(framing);
    let archive = converted_limits(archive, output)?;
    let insertions: Vec<EntryInsertion<'_>> = members
        .iter()
        .map(|member| EntryInsertion::append(&member.name, &member.data))
        .collect();
    Ok(catalog.reassemble_with_insertions_to_bytes(&insertions, archive)?)
}

fn walk(
    directory: &Path,
    prefix: &str,
    depth: usize,
    limits: CaptureLimits,
    members: &mut Vec<Member>,
    total: &mut u64,
) -> Result<(), CaptureError> {
    if depth > MAX_DEPTH {
        return Err(CaptureError::Invalid(
            "an iWork package directory nests deeper than 16 levels",
        ));
    }
    let metadata = fs::symlink_metadata(directory)?;
    if !metadata.is_dir() {
        return Err(CaptureError::Invalid(
            "an iWork package directory must be a real directory",
        ));
    }
    for entry in fs::read_dir(directory)? {
        let entry = entry?;
        let file_name = entry.file_name();
        let file_name = file_name.to_str().ok_or(CaptureError::Invalid(
            "an iWork package directory holds a non-UTF-8 name",
        ))?;
        let name = if prefix.is_empty() {
            file_name.to_owned()
        } else {
            format!("{prefix}/{file_name}")
        };
        let path = entry.path();
        let metadata = fs::symlink_metadata(&path)?;
        if metadata.file_type().is_symlink() {
            return Err(CaptureError::Invalid(
                "an iWork package directory must not contain symbolic links",
            ));
        }
        if metadata.is_dir() {
            walk(&path, &name, depth + 1, limits, members, total)?;
            continue;
        }
        if !metadata.is_file() {
            return Err(CaptureError::Invalid(
                "an iWork package directory must contain only regular files",
            ));
        }
        let count = u64::try_from(members.len() + 1).unwrap_or(u64::MAX);
        let maximum = u64::try_from(limits.max_members).unwrap_or(u64::MAX);
        if count > maximum {
            return Err(CaptureError::Limit {
                kind: CaptureLimitKind::Members,
                observed: count,
                maximum,
            });
        }
        let length = metadata.len();
        if length > limits.max_entry_bytes {
            return Err(CaptureError::Limit {
                kind: CaptureLimitKind::EntryBytes,
                observed: length,
                maximum: limits.max_entry_bytes,
            });
        }
        *total = total.saturating_add(length);
        if *total > limits.max_input_bytes {
            return Err(CaptureError::Limit {
                kind: CaptureLimitKind::InputBytes,
                observed: *total,
                maximum: limits.max_input_bytes,
            });
        }
        let amount = usize::try_from(length).unwrap_or(usize::MAX);
        let mut data = Vec::new();
        data.try_reserve_exact(amount)
            .map_err(|_error| CaptureError::Allocation { amount })?;
        fs::File::open(&path)?
            .take(length.saturating_add(1))
            .read_to_end(&mut data)?;
        if u64::try_from(data.len()).unwrap_or(u64::MAX) != length {
            return Err(CaptureError::Invalid(
                "an iWork package directory changed while it was being read",
            ));
        }
        members.push(Member { name, data });
    }
    Ok(())
}

/// The archive profile that admits a converted bundle of `length` bytes:
/// `limits` with its input ceiling raised to the converted length (never
/// above the hard ceiling [`Limits::MAX_INPUT_BYTES`]).
///
/// # Errors
///
/// An [`Error`] when the raised profile is not a valid [`Limits`].
pub fn converted_limits(limits: Limits, length: usize) -> crate::Result<Limits> {
    let length = u64::try_from(length).unwrap_or(u64::MAX);
    let input = limits
        .max_input_bytes()
        .max(length)
        .min(Limits::MAX_INPUT_BYTES);
    Limits::new(
        input,
        limits.max_entries(),
        limits.max_entry_bytes(),
        limits.max_total_bytes(),
        limits.max_iwa_stream_bytes(),
    )?
    .with_archive_limits(limits.archive_limits())
}

/// The bundle form of `package`: `Index.zip` (every `Index/` member, its
/// record copied unchanged) and every other member as a plain file.
fn bundle_members(
    package: &Catalog,
    limits: Limits,
) -> crate::Result<(Vec<u8>, Vec<(String, Vec<u8>)>)> {
    let mut others = Vec::new();
    let mut outside: Vec<&str> = Vec::new();
    for entry in package.iter() {
        if entry.name().starts_with("Index/") {
            continue;
        }
        let name = entry.name();
        if name == INDEX_ZIP
            || name.is_empty()
            || name.starts_with('/')
            || name.ends_with('/')
            || name
                .split('/')
                .any(|part| part.is_empty() || part == "." || part == "..")
        {
            return Err(Error::InvalidBundle(
                "a member name cannot be written as a bundle file".to_owned(),
            ));
        }
        outside.push(name);
        others.push((name.to_owned(), entry.data().to_vec()));
    }
    let index = package.reassemble_with_changes_to_bytes(&[], &[], &outside, limits)?;
    Ok((index, others))
}

/// Write the bundle into the fresh staging directory `root`; the
/// publication boundary synchronizes and swaps it into place.
fn write_bundle(root: &Path, index: &[u8], others: &[(String, Vec<u8>)]) -> io::Result<()> {
    use std::io::Write as _;
    let write_file = |path: &Path, data: &[u8]| -> io::Result<()> {
        fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(path)?
            .write_all(data)
    };
    write_file(&root.join(INDEX_ZIP), index)?;
    for (name, data) in others {
        let path = root.join(name);
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent)?;
        }
        write_file(&path, data)?;
    }
    Ok(())
}

/// Save `package` as a package directory at `path` through the directory
/// publication boundary (a sibling staging directory, synchronized, then
/// swapped in). `limits` bounds the `Index.zip` reassembly; `bundle` and
/// `bundling` classify an assembly failure and a staged-file write failure
/// in the caller's error type.
///
/// # Errors
///
/// A publication error (`E::from`) when the destination exists and is not a
/// directory or the staging, synchronization or swap fails; `bundling(..)`
/// when a staged file cannot be written; `bundle(..)` when the index archive
/// cannot be assembled or a member name cannot be a file.
pub fn save<E>(
    package: &Catalog,
    limits: Limits,
    path: &Path,
    bundle: impl FnOnce(Error) -> E,
    bundling: impl FnOnce(io::Error) -> E,
) -> Result<(), E>
where
    E: From<crate::publication::Error>,
{
    let (index, others) = bundle_members(package, limits).map_err(bundle)?;
    crate::publication::replace_directory_with::<E>(path, |root| {
        write_bundle(root, &index, &others).map_err(bundling)
    })
}
