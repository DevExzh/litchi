//! Package implementation for legacy `PowerPoint` presentations (`.ppt`).

pub(crate) mod codec;
pub(crate) mod model;
mod owner;
pub(crate) use model::Reader;
pub use owner::Package;
pub mod edit;
pub mod metadata;
#[allow(
    dead_code,
    reason = "the property-set owner behind package::metadata; its full surface stays exercised by the unit tests (record 0357)"
)]
pub(crate) mod property_set;
mod source;

#[cfg(test)]
mod tests;

pub use model::{
    EncryptionKind, Error, OpenOptions, Result, SourceLimitKind, SourceLimits, SourceLimitsBuilder,
};
pub use source::SourceBackedPackage;

/// Preflight all declared CFB stream lengths for a positional owner before a
/// format projection is allowed to materialize arbitrary metadata streams.
pub(crate) fn preflight_shared_ole_file(
    shared: &litchi_cfb::SharedOleFile,
    source_limits: SourceLimits,
) -> Result<u64> {
    // Fence both sides of the metadata-only scan. A stale positional source
    // must win over a limit result from a directory that is no longer current.
    shared.source_version()?;
    let result = preflight_input_bytes(shared.file_size(), source_limits).and_then(|()| {
        codec::preflight_directory_entries(shared.directory_entries(), source_limits)
    });
    shared.source_version()?;
    result
}

/// Preflight every declared stream in an already parsed eager CFB owner.
pub(crate) fn preflight_ole_file<R: std::io::Read + std::io::Seek>(
    ole: &litchi_cfb::OleFile<R>,
    source_limits: SourceLimits,
) -> Result<u64> {
    preflight_input_bytes(ole.file_size(), source_limits)?;
    codec::preflight_declared_streams(ole, source_limits)
}

fn preflight_input_bytes(bytes: u64, limits: SourceLimits) -> Result<()> {
    let observed = usize::try_from(bytes).unwrap_or(usize::MAX);
    let maximum = limits.max_input_bytes();
    if observed > maximum {
        return Err(Error::Limit {
            kind: SourceLimitKind::InputBytes,
            observed,
            maximum,
        });
    }
    Ok(())
}
