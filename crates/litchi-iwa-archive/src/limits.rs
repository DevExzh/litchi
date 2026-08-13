use litchi_iwa_core::{ArchiveLimits, SnappyLimits, SnappyStream};

use crate::{Error, LimitKind, Result};

/// Checked resource ceilings for one physical iWork bundle ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_input_bytes: u64,
    max_entries: usize,
    max_metadata_bytes: u64,
    max_entry_bytes: u64,
    max_total_bytes: u64,
    max_iwa_stream_bytes: usize,
    iwa_profile: ArchiveLimits,
}

impl Limits {
    /// Hard ceiling for bytes read from one bundle file or nested `Index.zip`.
    pub const MAX_INPUT_BYTES: u64 = 2 * 1024 * 1024 * 1024;
    /// Hard ceiling for physical ZIP members, including directory records.
    pub const MAX_ENTRIES: usize = 100_000;
    /// Hard ceiling for one raw ZIP member name.
    pub const MAX_MEMBER_NAME_BYTES: u64 = 4 * 1024;
    /// Hard ceiling for aggregate raw ZIP header variable metadata.
    pub const MAX_METADATA_BYTES: u64 = 64 * 1024 * 1024;
    /// Hard ceiling for one declared compressed ZIP member.
    pub const MAX_COMPRESSED_ENTRY_BYTES: u64 = 512 * 1024 * 1024;
    /// Hard ceiling for one declared uncompressed ZIP member.
    pub const MAX_ENTRY_BYTES: u64 = 512 * 1024 * 1024;
    /// Hard ceiling for one declared ZIP archive total.
    pub const MAX_TOTAL_BYTES: u64 = 2 * 1024 * 1024 * 1024;
    /// Hard ceiling for one decompressed IWA component.
    pub const MAX_IWA_STREAM_BYTES: usize = SnappyStream::MAX_DECOMPRESSED_STREAM;

    /// Build a checked physical ingress profile.
    ///
    /// # Errors
    ///
    /// Returns an error when any requested ceiling is zero or exceeds its
    /// format-wide hard ceiling.
    pub fn new(
        max_input_bytes: u64,
        max_entries: usize,
        max_entry_bytes: u64,
        max_total_bytes: u64,
        max_iwa_stream_bytes: usize,
    ) -> Result<Self> {
        let iwa_profile = ArchiveLimits::default()
            .with_archive_bytes(max_iwa_stream_bytes)
            .map_err(|error| Error::InvalidLimits(error.to_string()))?;
        Self {
            max_input_bytes,
            max_entries,
            max_metadata_bytes: max_input_bytes.min(Self::MAX_METADATA_BYTES),
            max_entry_bytes,
            max_total_bytes,
            max_iwa_stream_bytes,
            iwa_profile,
        }
        .validate()
    }

    /// Maximum complete input accepted from a path or byte slice.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }

    /// Maximum number of physical ZIP members, including directory records.
    #[must_use]
    pub const fn max_entries(self) -> usize {
        self.max_entries
    }

    /// Maximum aggregate logical or ZIP-header metadata bytes.
    pub(crate) const fn max_metadata_bytes(self) -> u64 {
        self.max_metadata_bytes
    }

    /// Set a derived subordinate metadata ceiling without weakening the outer
    /// checked profile.
    pub(crate) fn with_derived_metadata_bytes(mut self, maximum: u64) -> Result<Self> {
        if maximum > self.max_input_bytes.min(Self::MAX_METADATA_BYTES) {
            return Err(Error::Limit {
                kind: LimitKind::MetadataBytes,
                observed: maximum,
                maximum: self.max_input_bytes.min(Self::MAX_METADATA_BYTES),
            });
        }
        self.max_metadata_bytes = maximum;
        Ok(self)
    }

    /// Maximum declared uncompressed size of one ZIP member.
    #[must_use]
    pub const fn max_entry_bytes(self) -> u64 {
        self.max_entry_bytes
    }

    /// Maximum aggregate declared uncompressed ZIP size.
    #[must_use]
    pub const fn max_total_bytes(self) -> u64 {
        self.max_total_bytes
    }

    /// Maximum decompressed size of one IWA component.
    #[must_use]
    pub const fn max_iwa_stream_bytes(self) -> usize {
        self.max_iwa_stream_bytes
    }

    /// Return the caller-selected neutral IWA archive profile.
    #[must_use]
    pub const fn archive_limits(self) -> ArchiveLimits {
        self.iwa_profile
    }

    /// Tighten the neutral IWA archive profile while preserving the outer
    /// bundle stream ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error if the supplied neutral profile is invalid.
    pub fn with_archive_limits(mut self, limits: ArchiveLimits) -> Result<Self> {
        limits
            .validate()
            .map_err(|error| Error::InvalidLimits(error.to_string()))?;
        if limits.max_archive_bytes() > self.max_iwa_stream_bytes {
            return Err(Error::InvalidLimits(
                "IWA archive limit exceeds the bundle stream limit".to_owned(),
            ));
        }
        self.iwa_profile = limits;
        self.validate()
    }

    /// Return the effective neutral IWA profile applied during parsing.
    ///
    /// # Errors
    ///
    /// Returns an error if the profile cannot be tightened to the stream
    /// ceiling.
    pub fn effective_archive_limits(self) -> Result<ArchiveLimits> {
        Ok(self.validate()?.iwa_profile)
    }

    /// Return the checked Snappy framing profile corresponding to this budget.
    ///
    /// # Errors
    ///
    /// Returns an error if the effective IWA profile cannot be represented by
    /// the Snappy framing limits.
    pub fn snappy_limits(self) -> Result<SnappyLimits> {
        let max_stream_bytes = self.effective_archive_limits()?.max_archive_bytes();
        SnappyLimits::new(
            max_stream_bytes.min(SnappyStream::MAX_UNCOMPRESSED_CHUNK),
            max_stream_bytes,
        )
        .map_err(|error| Error::InvalidLimits(error.to_string()))
    }

    /// Reject an input whose complete byte length exceeds this profile.
    ///
    /// # Errors
    ///
    /// Returns an error when `size` exceeds the configured input ceiling.
    pub(crate) fn check_input_size(self, size: u64, _label: &str) -> Result<()> {
        if size > self.max_input_bytes {
            return Err(Error::Limit {
                kind: LimitKind::InputBytes,
                observed: size,
                maximum: self.max_input_bytes,
            });
        }
        Ok(())
    }

    /// Reject an output artifact before its complete byte buffer is allocated.
    ///
    /// The physical input ceiling is also the maximum in-memory artifact size
    /// for this bounded reassembly API. Callers that need a larger artifact
    /// must choose a larger profile, still subject to the format hard ceiling.
    pub(crate) fn check_output_size(self, size: u64) -> Result<()> {
        if size > self.max_input_bytes {
            return Err(Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: size,
                maximum: self.max_input_bytes,
            });
        }
        Ok(())
    }

    /// Charge decompressed IWA stream bytes before another parsed component is
    /// retained in the catalog.
    ///
    /// The aggregate reuses the package-wide uncompressed-byte ceiling rather
    /// than adding a second caller-facing knob. Individual components are
    /// still bounded independently by [`Self::max_iwa_stream_bytes`].
    pub(crate) fn charge_iwa_total_bytes(self, current: u64, added: u64) -> Result<u64> {
        let observed = current.saturating_add(added);
        if observed > self.max_total_bytes {
            return Err(Error::Limit {
                kind: LimitKind::IwaTotalBytes,
                observed,
                maximum: self.max_total_bytes,
            });
        }
        Ok(observed)
    }

    pub(crate) fn validate(self) -> Result<Self> {
        if self.max_input_bytes == 0
            || self.max_entries == 0
            || self.max_entry_bytes == 0
            || self.max_total_bytes == 0
            || self.max_iwa_stream_bytes == 0
        {
            return Err(Error::InvalidLimits(
                "all iWork archive limits must be non-zero".to_owned(),
            ));
        }
        if self.max_input_bytes > Self::MAX_INPUT_BYTES {
            return Err(Error::Limit {
                kind: LimitKind::InputBytes,
                observed: self.max_input_bytes,
                maximum: Self::MAX_INPUT_BYTES,
            });
        }
        if self.max_entries > Self::MAX_ENTRIES {
            return Err(Error::Limit {
                kind: LimitKind::Entries,
                observed: self.max_entries as u64,
                maximum: Self::MAX_ENTRIES as u64,
            });
        }
        if self.max_metadata_bytes > self.max_input_bytes.min(Self::MAX_METADATA_BYTES) {
            return Err(Error::Limit {
                kind: LimitKind::MetadataBytes,
                observed: self.max_metadata_bytes,
                maximum: self.max_input_bytes.min(Self::MAX_METADATA_BYTES),
            });
        }
        if self.max_entry_bytes > Self::MAX_ENTRY_BYTES {
            return Err(Error::Limit {
                kind: LimitKind::EntryBytes,
                observed: self.max_entry_bytes,
                maximum: Self::MAX_ENTRY_BYTES,
            });
        }
        if self.max_total_bytes > Self::MAX_TOTAL_BYTES {
            return Err(Error::Limit {
                kind: LimitKind::TotalBytes,
                observed: self.max_total_bytes,
                maximum: Self::MAX_TOTAL_BYTES,
            });
        }
        if self.max_iwa_stream_bytes > Self::MAX_IWA_STREAM_BYTES {
            return Err(Error::Limit {
                kind: LimitKind::IwaStreamBytes,
                observed: self.max_iwa_stream_bytes as u64,
                maximum: Self::MAX_IWA_STREAM_BYTES as u64,
            });
        }
        self.iwa_profile
            .validate()
            .map_err(|error| Error::InvalidLimits(error.to_string()))?;
        if self.iwa_profile.max_archive_bytes() > self.max_iwa_stream_bytes {
            return Err(Error::InvalidLimits(
                "IWA archive limit exceeds the bundle stream limit".to_owned(),
            ));
        }
        Ok(self)
    }

    pub(crate) const fn zip_limits(self) -> soapberry_zip::office::ArchiveLimits {
        soapberry_zip::office::ArchiveLimits {
            max_files: self.max_entries,
            max_member_name_bytes: if self.max_input_bytes < Self::MAX_MEMBER_NAME_BYTES {
                self.max_input_bytes
            } else {
                Self::MAX_MEMBER_NAME_BYTES
            },
            max_metadata_bytes: self.max_metadata_bytes,
            max_compressed_size: if self.max_input_bytes < Self::MAX_COMPRESSED_ENTRY_BYTES {
                self.max_input_bytes
            } else {
                Self::MAX_COMPRESSED_ENTRY_BYTES
            },
            max_entry_size: self.max_entry_bytes,
            max_total_size: self.max_total_bytes,
        }
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_input_bytes: Self::MAX_INPUT_BYTES,
            max_entries: Self::MAX_ENTRIES,
            max_metadata_bytes: Self::MAX_METADATA_BYTES,
            max_entry_bytes: Self::MAX_ENTRY_BYTES,
            max_total_bytes: Self::MAX_TOTAL_BYTES,
            max_iwa_stream_bytes: Self::MAX_IWA_STREAM_BYTES,
            iwa_profile: ArchiveLimits::default(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::Limits;

    #[test]
    fn zip_policy_maps_every_backend_resource_ceiling() {
        let defaults = Limits::default().zip_limits();
        assert_eq!(defaults.max_files, Limits::MAX_ENTRIES);
        assert_eq!(
            defaults.max_member_name_bytes,
            Limits::MAX_MEMBER_NAME_BYTES
        );
        assert_eq!(defaults.max_metadata_bytes, Limits::MAX_METADATA_BYTES);
        assert_eq!(
            defaults.max_compressed_size,
            Limits::MAX_COMPRESSED_ENTRY_BYTES
        );
        assert_eq!(defaults.max_entry_size, Limits::MAX_ENTRY_BYTES);
        assert_eq!(defaults.max_total_size, Limits::MAX_TOTAL_BYTES);

        let tight = Limits::new(7, 1, 1, 1, 1)
            .unwrap_or_else(|error| panic!("tight limits should be valid: {error}"))
            .zip_limits();
        assert_eq!(tight.max_member_name_bytes, 7);
        assert_eq!(tight.max_metadata_bytes, 7);
        assert_eq!(tight.max_compressed_size, 7);
    }
}
