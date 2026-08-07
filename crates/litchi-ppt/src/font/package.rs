use super::{FontCollections, Limits};
use crate::package::{Package, Result};
use litchi_cfb::OleFile;
use std::io::{Cursor, Read, Seek};

/// Whole-artifact and semantic limits captured by font snapshots and patches.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PackageLimits {
    pub max_source_bytes: usize,
    pub fonts: Limits,
}

impl Default for PackageLimits {
    fn default() -> Self {
        Self {
            max_source_bytes: 512 * 1024 * 1024,
            fonts: Limits::default(),
        }
    }
}

/// Borrowed open credentials plus finite package limits.
#[derive(Debug, Clone, Copy)]
pub struct PackageOptions<'a> {
    pub password: Option<&'a str>,
    pub limits: PackageLimits,
}

impl Default for PackageOptions<'_> {
    fn default() -> Self {
        Self {
            password: None,
            limits: PackageLimits::default(),
        }
    }
}

impl<R: Read + Seek> Package<R> {
    /// Read font semantics from the exact live persisted document.
    pub fn fonts(&mut self) -> Result<FontCollections> {
        self.presentation()?.fonts()
    }

    pub fn fonts_with_limits(&mut self, limits: Limits) -> Result<FontCollections> {
        self.presentation_with_limits(limits.records)?
            .fonts_with_limits(limits)
    }
}

/// Require every unrelated stream path and payload to remain exact.
pub(crate) fn validate_unrelated_streams(before: &[u8], after: &[u8]) -> Result<()> {
    let mut old = OleFile::open(Cursor::new(before))?;
    let mut new = OleFile::open(Cursor::new(after))?;
    let mut old_paths = old.list_streams();
    let mut new_paths = new.list_streams();
    old_paths.sort();
    new_paths.sort();
    if old_paths != new_paths {
        return Err(crate::package::Error::Corrupted(
            "font publication changed the CFB stream topology".into(),
        ));
    }
    for path in old_paths {
        if path.last().is_some_and(|name| {
            name.eq_ignore_ascii_case("PowerPoint Document")
                || name.eq_ignore_ascii_case("Current User")
        }) {
            continue;
        }
        let refs: Vec<_> = path.iter().map(String::as_str).collect();
        if old.open_stream(&refs)? != new.open_stream(&refs)? {
            return Err(crate::package::Error::Corrupted(format!(
                "font publication changed unrelated stream {}",
                path.join("/")
            )));
        }
    }
    Ok(())
}

/// Refuse nested CFB storages because the shared stream-only publisher cannot
/// prove preservation of their CLSIDs, state bits, timestamps, or emptiness.
pub(crate) fn require_stream_only_cfb(source: &[u8]) -> Result<()> {
    let ole = OleFile::open(Cursor::new(source))?;
    if ole
        .list_directory_entries(&[])?
        .iter()
        .any(|entry| entry.entry_type == 1)
    {
        return Err(crate::package::Error::InvalidFormat(
            "font editing refuses nested CFB storage metadata that cannot be preserved".into(),
        ));
    }
    Ok(())
}
