//! App-authored Pages package directories (IWORK_API_REQUESTS.md request 16).
//!
//! Pages can save a document as a directory bundle: `Index.zip` holds the
//! IWA components (already under `Index/`), and the metadata sidecars,
//! previews and `Data/` media sit beside it as plain files. The single-file
//! form is one ZIP holding the same members. [`single_file_bytes`] captures a
//! bundle once, under the caller's [`SourceLimits`], and converts it: the
//! `Index.zip` members first (their records copied as stored), then every
//! other file of the bundle as a stored member, in byte order of its relative
//! name, so the conversion is deterministic.
//!
//! The walk refuses symbolic links and special files, bounds the depth, the
//! member count and every byte figure, and re-checks each file's length after
//! reading it, so a bundle that changes while it is captured is refused
//! rather than half-read.

use std::fs;
use std::io::Read;
use std::path::Path;

use litchi_iwa_archive::package::{Catalog, EntryInsertion};

use super::{PackageError, PackageResult};
use crate::{LimitKind, SourceLimitKind, SourceLimits};

/// The deepest directory nesting a bundle may use below its root.
const MAX_DEPTH: usize = 16;
/// The IWA component archive at the root of a bundle.
const INDEX_ZIP: &str = "Index.zip";

/// One captured bundle file: its `/`-separated relative name and bytes.
struct Member {
    name: String,
    data: Vec<u8>,
}

fn limit(kind: SourceLimitKind, observed: u64, maximum: u64) -> PackageError {
    PackageError::Limit {
        kind: LimitKind::Source(kind),
        observed,
        maximum,
    }
}

fn invalid(message: &str) -> PackageError {
    PackageError::InvalidFormat(message.to_owned())
}

/// Capture the bundle at `root` and return its single-file package bytes.
///
/// # Errors
///
/// [`PackageError::InvalidFormat`] for a bundle without `Index.zip`, with a
/// loose `Index/` directory, a symbolic link, a special file, a non-UTF-8
/// name or nesting deeper than 16 levels; [`PackageError::Limit`] when a
/// figure exceeds `limits`; [`PackageError::Io`] for a read failure;
/// archive errors when `Index.zip` is malformed or the reassembly fails.
pub(super) fn single_file_bytes(root: &Path, limits: SourceLimits) -> PackageResult<Vec<u8>> {
    let mut members = Vec::new();
    let mut total = 0_u64;
    walk(root, "", 0, limits, &mut members, &mut total)?;
    members.sort_by(|left, right| left.name.as_bytes().cmp(right.name.as_bytes()));
    let index = members
        .iter()
        .position(|member| member.name == INDEX_ZIP)
        .ok_or_else(|| invalid("a Pages package directory has no Index.zip"))?;
    if members
        .iter()
        .any(|member| member.name.starts_with("Index/"))
    {
        return Err(invalid(
            "Pages package directories with a loose Index folder are not supported",
        ));
    }
    let index = members.remove(index);
    let archive_limits = limits.archive()?;
    let catalog = Catalog::from_bytes_with_limits(&index.data, archive_limits)?;
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
    let archive_limits = converted_limits(archive_limits, output)?;
    let insertions: Vec<EntryInsertion<'_>> = members
        .iter()
        .map(|member| EntryInsertion::append(&member.name, &member.data))
        .collect();
    Ok(catalog.reassemble_with_insertions_to_bytes(&insertions, archive_limits)?)
}

fn walk(
    directory: &Path,
    prefix: &str,
    depth: usize,
    limits: SourceLimits,
    members: &mut Vec<Member>,
    total: &mut u64,
) -> PackageResult<()> {
    if depth > MAX_DEPTH {
        return Err(invalid(
            "a Pages package directory nests deeper than 16 levels",
        ));
    }
    let metadata = fs::symlink_metadata(directory)?;
    if !metadata.is_dir() {
        return Err(invalid(
            "a Pages package directory must be a real directory",
        ));
    }
    for entry in fs::read_dir(directory)? {
        let entry = entry?;
        let file_name = entry.file_name();
        let file_name = file_name
            .to_str()
            .ok_or_else(|| invalid("a Pages package directory holds a non-UTF-8 name"))?;
        let name = if prefix.is_empty() {
            file_name.to_owned()
        } else {
            format!("{prefix}/{file_name}")
        };
        let path = entry.path();
        let metadata = fs::symlink_metadata(&path)?;
        if metadata.file_type().is_symlink() {
            return Err(invalid(
                "a Pages package directory must not contain symbolic links",
            ));
        }
        if metadata.is_dir() {
            walk(&path, &name, depth + 1, limits, members, total)?;
            continue;
        }
        if !metadata.is_file() {
            return Err(invalid(
                "a Pages package directory must contain only regular files",
            ));
        }
        let count = u64::try_from(members.len() + 1).unwrap_or(u64::MAX);
        let maximum = u64::try_from(limits.max_members()).unwrap_or(u64::MAX);
        if count > maximum {
            return Err(limit(SourceLimitKind::Members, count, maximum));
        }
        let length = metadata.len();
        if length > limits.max_entry_bytes() {
            return Err(limit(
                SourceLimitKind::EntryBytes,
                length,
                limits.max_entry_bytes(),
            ));
        }
        *total = total.saturating_add(length);
        let input = limits.max_input_bytes().min(limits.max_aggregate_bytes());
        if *total > input {
            return Err(limit(SourceLimitKind::InputBytes, *total, input));
        }
        let mut data = Vec::new();
        data.try_reserve_exact(usize::try_from(length).unwrap_or(usize::MAX))
            .map_err(|_error| PackageError::Allocation {
                amount: usize::try_from(length).unwrap_or(usize::MAX),
            })?;
        fs::File::open(&path)?
            .take(length.saturating_add(1))
            .read_to_end(&mut data)?;
        if u64::try_from(data.len()).unwrap_or(u64::MAX) != length {
            return Err(invalid(
                "a Pages package directory changed while it was being read",
            ));
        }
        members.push(Member { name, data });
    }
    Ok(())
}

/// The archive profile that admits a converted bundle of `length` bytes:
/// `limits` with its input ceiling raised to the converted length (never
/// above the hard ceiling).
pub(super) fn converted_limits(
    limits: litchi_iwa_archive::Limits,
    length: usize,
) -> PackageResult<litchi_iwa_archive::Limits> {
    let length = u64::try_from(length).unwrap_or(u64::MAX);
    let input = limits
        .max_input_bytes()
        .max(length)
        .min(litchi_iwa_archive::Limits::MAX_INPUT_BYTES);
    Ok(litchi_iwa_archive::Limits::new(
        input,
        limits.max_entries(),
        limits.max_entry_bytes(),
        limits.max_total_bytes(),
        limits.max_iwa_stream_bytes(),
    )?
    .with_archive_limits(limits.archive_limits())?)
}
