use litchi_iwa_core::{Archive, SnappyStream};
use soapberry_zip::office::ArchiveReader;

use crate::catalog::Component;
use crate::{Error, Limits, Result};

/// Opaque ZIP reader used by the physical component catalog.
pub(crate) struct ZipArchive<'data> {
    reader: ArchiveReader<'data>,
}

impl<'data> ZipArchive<'data> {
    pub(crate) fn new_with_limits(data: &'data [u8], limits: Limits) -> Result<Self> {
        let validated_limits = limits.validate()?;
        let input_size = u64::try_from(data.len()).map_err(|_error| {
            Error::InvalidBundle("ZIP input length does not fit u64".to_owned())
        })?;
        validated_limits.check_input_size(input_size, "ZIP input")?;
        Ok(Self {
            reader: ArchiveReader::new_with_limits(data, validated_limits.zip_limits())?,
        })
    }

    pub(crate) fn file_names(&self) -> impl Iterator<Item = &str> {
        self.reader.file_names()
    }

    pub(crate) fn read(&self, name: &str) -> Result<Vec<u8>> {
        Ok(self.reader.read(name)?)
    }
}

pub(crate) fn parse_iwa_components(
    archive: &ZipArchive<'_>,
    limits: Limits,
) -> Result<Vec<Component>> {
    let validated_limits = limits.validate()?;
    if is_encrypted(archive) {
        return Err(Error::Encrypted);
    }

    if archive.file_names().any(is_iwa_name) {
        return parse_direct_iwa_components(archive, validated_limits);
    }

    let Some(index_name) = nested_index_name(archive)? else {
        return Ok(Vec::new());
    };
    let index_data = archive.read(&index_name)?;
    let index_size = u64::try_from(index_data.len()).map_err(|_error| {
        Error::InvalidBundle("legacy iWork Index.zip length does not fit u64".to_owned())
    })?;
    validated_limits.check_input_size(index_size, "legacy iWork Index.zip")?;
    let index = ZipArchive::new_with_limits(&index_data, validated_limits)?;
    let components = parse_direct_iwa_components(&index, validated_limits)?;
    if components.is_empty() {
        return Err(Error::InvalidBundle(format!(
            "legacy package index {index_name} contains no IWA components"
        )));
    }
    Ok(components)
}

fn parse_direct_iwa_components(archive: &ZipArchive<'_>, limits: Limits) -> Result<Vec<Component>> {
    let mut components = Vec::new();
    for name in archive.file_names() {
        if !is_iwa_name(name) {
            continue;
        }
        let compressed_data = archive.read(name)?;

        // OperationStorage is a separate persistence format despite its `.iwa`
        // suffix in legacy documents. It remains a raw package member but is
        // not part of the IWA object graph.
        if name.rsplit('/').next() == Some("OperationStorage.iwa")
            && compressed_data.starts_with(b"bvxn")
        {
            continue;
        }

        let decompressed =
            SnappyStream::decompress_with_limits(&compressed_data, limits.snappy_limits()?)?;
        let parsed = Archive::parse_with_limits(
            decompressed.as_bytes(),
            limits.effective_archive_limits()?,
        )?;
        components.push(Component::new(name, parsed));
    }
    components.sort_unstable_by(|left, right| left.name().cmp(right.name()));
    Ok(components)
}

pub(crate) fn is_encrypted(archive: &ZipArchive<'_>) -> bool {
    archive
        .file_names()
        .any(|name| matches!(name.rsplit('/').next(), Some(".iwpv2" | ".iwph")))
}

pub(crate) fn nested_index_name(archive: &ZipArchive<'_>) -> Result<Option<String>> {
    let mut candidates = archive
        .file_names()
        .filter(|name| name.rsplit('/').next() == Some("Index.zip"));
    let first = candidates.next().map(str::to_owned);
    if let Some(second) = candidates.next() {
        return Err(Error::InvalidBundle(format!(
            "iWork package contains ambiguous nested indexes: {} and {second}",
            first.as_deref().unwrap_or("Index.zip")
        )));
    }
    Ok(first)
}

pub(crate) fn is_iwa_name(name: &str) -> bool {
    #[allow(
        clippy::case_sensitive_file_extension_comparisons,
        reason = "IWA member names are case-sensitive protocol names."
    )]
    {
        name.ends_with(".iwa")
    }
}
