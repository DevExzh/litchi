//! Semantic state for the writer-side File Information Block.

use super::flags::BaseFlags;
use super::offsets::Offsets;

/// FIB builder state.
///
/// The public builder remains intentionally small: the serialized header,
/// story counts, and table references are kept in separate semantic owners so
/// each codec section can evolve independently.
#[derive(Debug)]
#[allow(
    dead_code,
    reason = "reserved DOC structure retained for format completeness or future round-trip support"
)]
pub struct FibBuilder {
    pub(super) header: Header,
    pub(super) stories: Stories,
    pub(super) offsets: Offsets,
}

/// Fields belonging to `FibBase`.
#[derive(Debug)]
#[allow(
    dead_code,
    reason = "reserved DOC structure retained for format completeness or future round-trip support"
)]
pub(super) struct Header {
    pub(super) text_size: u32,
    pub(super) table_size: u32,
    pub(super) flags: BaseFlags,
    pub(super) next_fib_page: u16,
    pub(super) fc_min: u32,
    pub(super) fc_mac: u32,
    pub(super) cb_mac: u32,
}

/// Character counts and story ranges emitted in `FibRgLw97`.
#[derive(Debug, Default)]
#[allow(
    dead_code,
    reason = "reserved DOC structure retained for format completeness or future round-trip support"
)]
pub(super) struct Stories {
    pub(super) main_text_start: u32,
    pub(super) main_text_length: u32,
    pub(super) footnote_start: u32,
    pub(super) footnote_length: u32,
    pub(super) header_start: u32,
    pub(super) header_length: u32,
    pub(super) comment_start: u32,
    pub(super) comment_length: u32,
    pub(super) endnote_start: u32,
    pub(super) endnote_length: u32,
    pub(super) textbox_start: u32,
    pub(super) textbox_length: u32,
    pub(super) header_textbox_length: u32,
}

impl FibBuilder {
    /// Create a new FIB builder.
    #[must_use]
    pub fn new() -> Self {
        Self {
            header: Header {
                text_size: 0,
                table_size: 0,
                flags: BaseFlags::default(),
                next_fib_page: 0,
                fc_min: 0,
                fc_mac: 0,
                cb_mac: 0,
            },
            stories: Stories::default(),
            offsets: Offsets::default(),
        }
    }

    /// Set the main document text range.
    pub fn set_main_text(&mut self, start: u32, length: u32) {
        self.stories.main_text_start = start;
        self.stories.main_text_length = length;
        self.header.text_size = start + length;
    }

    /// Set the table stream size.
    pub fn set_table_size(&mut self, size: u32) {
        self.header.table_size = size;
    }

    /// Set the main document base offsets and stream size.
    pub fn set_base_fields(&mut self, fc_min: u32, fc_mac: u32, cb_mac: u32) {
        self.header.fc_min = fc_min;
        self.header.fc_mac = fc_mac;
        self.header.cb_mac = cb_mac;
    }

    /// Mark this as a glossary-only FIB.
    pub fn set_glossary_document(&mut self, is_glossary: bool) {
        self.header.flags.glossary = is_glossary;
    }

    /// Mark this as a template FIB.
    pub fn set_template(&mut self, is_template: bool) {
        self.header.flags.template = is_template;
    }

    /// Address a secondary FIB by its 512-byte page number.
    pub fn set_next_fib_page(&mut self, page: u16) {
        self.header.next_fib_page = page;
    }

    /// Set the header/footer story character count.
    pub fn set_ccp_hdd(&mut self, length: u32) {
        self.stories.header_length = length;
    }

    /// Set the footnote story character count.
    pub fn set_ccp_ftn(&mut self, length: u32) {
        self.stories.footnote_length = length;
    }

    /// Set the endnote story character count.
    pub fn set_ccp_edn(&mut self, length: u32) {
        self.stories.endnote_length = length;
    }

    /// Set the comment story character count.
    pub fn set_ccp_atn(&mut self, length: u32) {
        self.stories.comment_length = length;
    }

    /// Set the textbox story character count.
    pub fn set_ccp_txbx(&mut self, length: u32) {
        self.stories.textbox_length = length;
    }

    /// Set the header textbox story character count.
    pub fn set_ccp_hdr_txbx(&mut self, length: u32) {
        self.stories.header_textbox_length = length;
    }
}

impl Default for FibBuilder {
    fn default() -> Self {
        Self::new()
    }
}
