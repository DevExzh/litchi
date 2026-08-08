use super::{Error, Resource, Result};

/// Physical resource ceilings used before format-owned semantic decoding.
#[allow(
    clippy::struct_field_names,
    reason = "The max_* vocabulary makes each independent ceiling explicit."
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SourceLimits {
    max_input_bytes: u64,
    max_entries: usize,
    max_entry_bytes: u64,
    max_expanded_bytes: u64,
    max_decoded_bytes_per_item: usize,
}

impl SourceLimits {
    /// Hard maximum for complete input bytes.
    pub const HARD_MAX_INPUT_BYTES: u64 = litchi_iwa_detect::Limits::HARD_MAX_INPUT_BYTES;
    /// Hard maximum for packaged entry count.
    pub const HARD_MAX_ENTRIES: usize = litchi_iwa_detect::Limits::HARD_MAX_FILES;
    /// Hard maximum for one expanded packaged entry.
    pub const HARD_MAX_ENTRY_BYTES: u64 = litchi_iwa_detect::Limits::HARD_MAX_ENTRY_SIZE;
    /// Hard maximum for aggregate expanded package bytes.
    pub const HARD_MAX_EXPANDED_BYTES: u64 = litchi_iwa_detect::Limits::HARD_MAX_TOTAL_SIZE;
    /// Hard maximum for one decoded package unit.
    pub const HARD_MAX_DECODED_BYTES_PER_ITEM: usize =
        litchi_iwa_detect::Limits::HARD_MAX_IWA_STREAM_SIZE;

    /// Construct a checked physical resource profile.
    ///
    /// # Errors
    ///
    /// Returns [`ErrorKind::InvalidOptions`](super::ErrorKind::InvalidOptions)
    /// when a ceiling is zero or exceeds its hard maximum.
    pub const fn new(
        max_input_bytes: u64,
        max_entries: usize,
        max_entry_bytes: u64,
        max_expanded_bytes: u64,
        max_decoded_bytes_per_item: usize,
    ) -> Result<Self> {
        if max_input_bytes == 0 || max_input_bytes > Self::HARD_MAX_INPUT_BYTES {
            return Err(Error::invalid_options(
                Resource::InputBytes,
                max_input_bytes,
                Self::HARD_MAX_INPUT_BYTES,
            ));
        }
        if max_entries == 0 || max_entries > Self::HARD_MAX_ENTRIES {
            return Err(Error::invalid_options(
                Resource::Entries,
                max_entries as u64,
                Self::HARD_MAX_ENTRIES as u64,
            ));
        }
        if max_entry_bytes == 0 || max_entry_bytes > Self::HARD_MAX_ENTRY_BYTES {
            return Err(Error::invalid_options(
                Resource::EntryBytes,
                max_entry_bytes,
                Self::HARD_MAX_ENTRY_BYTES,
            ));
        }
        if max_expanded_bytes == 0 || max_expanded_bytes > Self::HARD_MAX_EXPANDED_BYTES {
            return Err(Error::invalid_options(
                Resource::ExpandedBytes,
                max_expanded_bytes,
                Self::HARD_MAX_EXPANDED_BYTES,
            ));
        }
        if max_decoded_bytes_per_item == 0
            || max_decoded_bytes_per_item > Self::HARD_MAX_DECODED_BYTES_PER_ITEM
        {
            return Err(Error::invalid_options(
                Resource::DecodedBytes,
                max_decoded_bytes_per_item as u64,
                Self::HARD_MAX_DECODED_BYTES_PER_ITEM as u64,
            ));
        }
        Ok(Self {
            max_input_bytes,
            max_entries,
            max_entry_bytes,
            max_expanded_bytes,
            max_decoded_bytes_per_item,
        })
    }

    /// Maximum complete input bytes.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }

    /// Maximum packaged entry count.
    #[must_use]
    pub const fn max_entries(self) -> usize {
        self.max_entries
    }

    /// Maximum expanded bytes for one packaged entry.
    #[must_use]
    pub const fn max_entry_bytes(self) -> u64 {
        self.max_entry_bytes
    }

    /// Maximum aggregate expanded package bytes.
    #[must_use]
    pub const fn max_expanded_bytes(self) -> u64 {
        self.max_expanded_bytes
    }

    /// Maximum decoded bytes for one internal package unit.
    #[must_use]
    pub const fn max_decoded_bytes_per_item(self) -> usize {
        self.max_decoded_bytes_per_item
    }

    pub(super) fn detector(self) -> Result<litchi_iwa_detect::Limits> {
        litchi_iwa_detect::Limits::new(
            self.max_input_bytes,
            self.max_entries,
            self.max_entry_bytes,
            self.max_expanded_bytes,
            self.max_decoded_bytes_per_item,
        )
        .map_err(|_error| Error::invariant(None, super::Stage::Validation))
    }
}

impl Default for SourceLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: Self::HARD_MAX_INPUT_BYTES,
            max_entries: Self::HARD_MAX_ENTRIES,
            max_entry_bytes: Self::HARD_MAX_ENTRY_BYTES,
            max_expanded_bytes: Self::HARD_MAX_EXPANDED_BYTES,
            max_decoded_bytes_per_item: Self::HARD_MAX_DECODED_BYTES_PER_ITEM,
        }
    }
}

/// Archive-free semantic resource ceilings for one format-neutral snapshot.
#[allow(
    clippy::struct_field_names,
    reason = "The max_* vocabulary makes each independent ceiling explicit."
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SnapshotLimits {
    max_tables: usize,
    max_slides: usize,
    max_sections: usize,
    max_text_bytes: usize,
}

impl SnapshotLimits {
    /// Hard maximum for retained Numbers tables.
    pub const HARD_MAX_TABLES: usize = litchi_iwa_structured::MAX_TABLES;
    /// Hard maximum for retained Keynote slides.
    pub const HARD_MAX_SLIDES: usize = litchi_iwa_structured::MAX_SLIDES;
    /// Hard maximum for retained Pages sections.
    pub const HARD_MAX_SECTIONS: usize = litchi_iwa_structured::MAX_SECTIONS;
    /// Hard maximum for aggregate retained UTF-8 text.
    pub const HARD_MAX_TEXT_BYTES: usize = litchi_iwa_structured::DEFAULT_MAX_TEXT_BYTES;

    /// Construct a checked semantic resource profile.
    ///
    /// # Errors
    ///
    /// Returns [`ErrorKind::InvalidOptions`](super::ErrorKind::InvalidOptions)
    /// when a ceiling is zero or exceeds its hard maximum.
    pub const fn new(
        max_tables: usize,
        max_slides: usize,
        max_sections: usize,
        max_text_bytes: usize,
    ) -> Result<Self> {
        if max_tables == 0 || max_tables > Self::HARD_MAX_TABLES {
            return Err(Error::invalid_options(
                Resource::Tables,
                max_tables as u64,
                Self::HARD_MAX_TABLES as u64,
            ));
        }
        if max_slides == 0 || max_slides > Self::HARD_MAX_SLIDES {
            return Err(Error::invalid_options(
                Resource::Slides,
                max_slides as u64,
                Self::HARD_MAX_SLIDES as u64,
            ));
        }
        if max_sections == 0 || max_sections > Self::HARD_MAX_SECTIONS {
            return Err(Error::invalid_options(
                Resource::Sections,
                max_sections as u64,
                Self::HARD_MAX_SECTIONS as u64,
            ));
        }
        if max_text_bytes == 0 || max_text_bytes > Self::HARD_MAX_TEXT_BYTES {
            return Err(Error::invalid_options(
                Resource::TextBytes,
                max_text_bytes as u64,
                Self::HARD_MAX_TEXT_BYTES as u64,
            ));
        }
        Ok(Self {
            max_tables,
            max_slides,
            max_sections,
            max_text_bytes,
        })
    }

    /// Maximum retained Numbers tables.
    #[must_use]
    pub const fn max_tables(self) -> usize {
        self.max_tables
    }

    /// Maximum retained Keynote slides.
    #[must_use]
    pub const fn max_slides(self) -> usize {
        self.max_slides
    }

    /// Maximum retained Pages sections.
    #[must_use]
    pub const fn max_sections(self) -> usize {
        self.max_sections
    }

    /// Maximum aggregate retained UTF-8 text bytes.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    pub(super) fn aggregate(self) -> Result<litchi_iwa_structured::Limits> {
        litchi_iwa_structured::Limits::try_new(
            self.max_tables,
            self.max_slides,
            self.max_sections,
            self.max_text_bytes,
        )
        .map_err(|_error| Error::invariant(None, super::Stage::Validation))
    }
}

impl Default for SnapshotLimits {
    fn default() -> Self {
        Self {
            max_tables: Self::HARD_MAX_TABLES,
            max_slides: Self::HARD_MAX_SLIDES,
            max_sections: Self::HARD_MAX_SECTIONS,
            max_text_bytes: Self::HARD_MAX_TEXT_BYTES,
        }
    }
}

/// Complete checked options for one format-neutral iWork read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Options {
    source: SourceLimits,
    snapshot: SnapshotLimits,
}

impl Options {
    /// Combine independently checked physical and semantic profiles.
    #[must_use]
    pub const fn new(source: SourceLimits, snapshot: SnapshotLimits) -> Self {
        Self { source, snapshot }
    }

    /// Return the physical ingress profile.
    #[must_use]
    pub const fn source(self) -> SourceLimits {
        self.source
    }

    /// Return the archive-free semantic profile.
    #[must_use]
    pub const fn snapshot(self) -> SnapshotLimits {
        self.snapshot
    }

    /// Replace the physical ingress profile.
    #[must_use]
    pub const fn with_source(mut self, value: SourceLimits) -> Self {
        self.source = value;
        self
    }

    /// Replace the archive-free semantic profile.
    #[must_use]
    pub const fn with_snapshot(mut self, value: SnapshotLimits) -> Self {
        self.snapshot = value;
        self
    }
}
