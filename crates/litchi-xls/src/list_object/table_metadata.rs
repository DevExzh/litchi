//! Typed `TableFeatureType` metadata shared by the semantic and wire layers.
//!
//! The bit positions are defined by [MS-XLS] 2.5.266.  Unknown bits are kept
//! separately so a read/edit/write pass does not normalize flags that this
//! crate does not yet interpret.

const KNOWN_BITS: u32 = 0x017f_fb7e;

/// Flags from the fixed `TableFeatureType` metadata word.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TableFlags {
    auto_filter: bool,
    persist_auto_filter: bool,
    show_insert_row: bool,
    insert_row_inserts_cells: bool,
    load_deleted_row_ids: bool,
    shown_total_row: bool,
    needs_commit: bool,
    single_cell: bool,
    apply_auto_filter: bool,
    force_insert_to_be_visible: bool,
    compressed_xml: bool,
    load_provider_name: bool,
    load_changed_row_ids: bool,
    version_nibble: u8,
    load_entry_id: bool,
    load_invalid_cells: bool,
    good_build: bool,
    published: bool,
    unknown_bits: u32,
}

impl TableFlags {
    /// The canonical flags for a regular headered table with `AutoFilter` data.
    pub const fn default_table() -> Self {
        Self::from_raw(0x001B_0806)
    }

    /// Decode the wire flag word without rejecting future or reserved bits.
    pub const fn from_raw(raw: u32) -> Self {
        Self {
            auto_filter: raw & (1 << 1) != 0,
            persist_auto_filter: raw & (1 << 2) != 0,
            show_insert_row: raw & (1 << 3) != 0,
            insert_row_inserts_cells: raw & (1 << 4) != 0,
            load_deleted_row_ids: raw & (1 << 5) != 0,
            shown_total_row: raw & (1 << 6) != 0,
            needs_commit: raw & (1 << 8) != 0,
            single_cell: raw & (1 << 9) != 0,
            apply_auto_filter: raw & (1 << 11) != 0,
            force_insert_to_be_visible: raw & (1 << 12) != 0,
            compressed_xml: raw & (1 << 13) != 0,
            load_provider_name: raw & (1 << 14) != 0,
            load_changed_row_ids: raw & (1 << 15) != 0,
            version_nibble: ((raw >> 16) & 0xF) as u8,
            load_entry_id: raw & (1 << 20) != 0,
            load_invalid_cells: raw & (1 << 21) != 0,
            good_build: raw & (1 << 22) != 0,
            published: raw & (1 << 24) != 0,
            unknown_bits: raw & !KNOWN_BITS,
        }
    }

    /// Encode the exact flag word, including retained unknown bits.
    pub const fn raw(self) -> u32 {
        let mut raw = self.unknown_bits;
        raw |= (self.auto_filter as u32) << 1;
        raw |= (self.persist_auto_filter as u32) << 2;
        raw |= (self.show_insert_row as u32) << 3;
        raw |= (self.insert_row_inserts_cells as u32) << 4;
        raw |= (self.load_deleted_row_ids as u32) << 5;
        raw |= (self.shown_total_row as u32) << 6;
        raw |= (self.needs_commit as u32) << 8;
        raw |= (self.single_cell as u32) << 9;
        raw |= (self.apply_auto_filter as u32) << 11;
        raw |= (self.force_insert_to_be_visible as u32) << 12;
        raw |= (self.compressed_xml as u32) << 13;
        raw |= (self.load_provider_name as u32) << 14;
        raw |= (self.load_changed_row_ids as u32) << 15;
        raw |= (self.version_nibble as u32 & 0xF) << 16;
        raw |= (self.load_entry_id as u32) << 20;
        raw |= (self.load_invalid_cells as u32) << 21;
        raw |= (self.good_build as u32) << 22;
        raw |= (self.published as u32) << 24;
        raw
    }

    pub const fn auto_filter(self) -> bool {
        self.auto_filter
    }

    pub const fn persists_auto_filter(self) -> bool {
        self.persist_auto_filter
    }

    pub const fn shows_insert_row(self) -> bool {
        self.show_insert_row
    }

    pub const fn insert_row_inserts_cells(self) -> bool {
        self.insert_row_inserts_cells
    }

    pub const fn loads_deleted_row_ids(self) -> bool {
        self.load_deleted_row_ids
    }

    pub const fn shows_total_row(self) -> bool {
        self.shown_total_row
    }

    pub const fn needs_commit(self) -> bool {
        self.needs_commit
    }

    pub const fn is_single_cell(self) -> bool {
        self.single_cell
    }

    pub const fn applies_auto_filter(self) -> bool {
        self.apply_auto_filter
    }

    pub const fn forces_insert_row_visible(self) -> bool {
        self.force_insert_to_be_visible
    }

    pub const fn uses_compressed_xml(self) -> bool {
        self.compressed_xml
    }

    pub const fn loads_provider_name(self) -> bool {
        self.load_provider_name
    }

    pub const fn loads_changed_row_ids(self) -> bool {
        self.load_changed_row_ids
    }

    /// The four-bit `verXL` value from [MS-XLS] `TableFeatureType`.
    pub const fn version_nibble(self) -> u8 {
        self.version_nibble
    }

    pub const fn loads_entry_id(self) -> bool {
        self.load_entry_id
    }

    pub const fn loads_invalid_cells(self) -> bool {
        self.load_invalid_cells
    }

    pub const fn has_good_build(self) -> bool {
        self.good_build
    }

    pub const fn is_published(self) -> bool {
        self.published
    }

    pub const fn unknown_bits(self) -> u32 {
        self.unknown_bits
    }

    pub const fn with_auto_filter(mut self, value: bool) -> Self {
        self.auto_filter = value;
        self
    }

    pub const fn with_persist_auto_filter(mut self, value: bool) -> Self {
        self.persist_auto_filter = value;
        self
    }

    pub const fn with_show_insert_row(mut self, value: bool) -> Self {
        self.show_insert_row = value;
        self
    }

    pub const fn with_insert_row_inserts_cells(mut self, value: bool) -> Self {
        self.insert_row_inserts_cells = value;
        self
    }

    pub const fn with_load_deleted_row_ids(mut self, value: bool) -> Self {
        self.load_deleted_row_ids = value;
        self
    }

    pub const fn with_shown_total_row(mut self, value: bool) -> Self {
        self.shown_total_row = value;
        self
    }

    pub const fn with_needs_commit(mut self, value: bool) -> Self {
        self.needs_commit = value;
        self
    }

    pub const fn with_single_cell(mut self, value: bool) -> Self {
        self.single_cell = value;
        self
    }

    pub const fn with_apply_auto_filter(mut self, value: bool) -> Self {
        self.apply_auto_filter = value;
        self
    }

    pub const fn with_force_insert_to_be_visible(mut self, value: bool) -> Self {
        self.force_insert_to_be_visible = value;
        self
    }

    pub const fn with_compressed_xml(mut self, value: bool) -> Self {
        self.compressed_xml = value;
        self
    }

    pub const fn with_load_provider_name(mut self, value: bool) -> Self {
        self.load_provider_name = value;
        self
    }

    pub const fn with_load_changed_row_ids(mut self, value: bool) -> Self {
        self.load_changed_row_ids = value;
        self
    }

    pub const fn with_version_nibble(mut self, value: u8) -> Self {
        self.version_nibble = value & 0xF;
        self
    }

    pub const fn with_load_entry_id(mut self, value: bool) -> Self {
        self.load_entry_id = value;
        self
    }

    pub const fn with_load_invalid_cells(mut self, value: bool) -> Self {
        self.load_invalid_cells = value;
        self
    }

    pub const fn with_good_build(mut self, value: bool) -> Self {
        self.good_build = value;
        self
    }

    pub const fn with_published(mut self, value: bool) -> Self {
        self.published = value;
        self
    }

    /// Preserve a caller-provided future/reserved bit mask.
    pub const fn with_unknown_bits(mut self, value: u32) -> Self {
        self.unknown_bits = value & !KNOWN_BITS;
        self
    }
}
