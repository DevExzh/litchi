/// Status of the last abstract-numbering cleanup attempt.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AbstractNumberingCleanupStatus {
    /// All abstract numbering definitions are considered reviewed (`0`).
    Reviewed,
    /// The last cleanup attempt was incomplete (`1`).
    Incomplete,
}

/// Passive Word document-event mask from `grfdoceventsN`.
///
/// Bits 6 and 7 are documented as reserved for internal use. They are retained
/// but never interpreted or executed by this crate.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentEventMask(u16);

impl DocumentEventMask {
    pub const NEW: Self = Self(1 << 0);
    pub const OPEN: Self = Self(1 << 1);
    pub const CLOSE: Self = Self(1 << 2);
    pub const SYNC: Self = Self(1 << 3);
    pub const XML_AFTER_INSERT: Self = Self(1 << 4);
    pub const XML_BEFORE_DELETE: Self = Self(1 << 5);
    pub const RESERVED_INTERNAL_6: Self = Self(1 << 6);
    pub const RESERVED_INTERNAL_7: Self = Self(1 << 7);
    pub const CONTENT_CONTROL_AFTER_ADD: Self = Self(1 << 8);
    pub const CONTENT_CONTROL_BEFORE_DELETE: Self = Self(1 << 9);
    pub const CONTENT_CONTROL_ON_EXIT: Self = Self(1 << 10);
    pub const CONTENT_CONTROL_ON_ENTER: Self = Self(1 << 11);
    pub const CONTENT_CONTROL_BEFORE_STORE_UPDATE: Self = Self(1 << 12);
    pub const CONTENT_CONTROL_BEFORE_CONTENT_UPDATE: Self = Self(1 << 13);
    pub const BUILDING_BLOCK_INSERT: Self = Self(1 << 14);
    pub const ALL: Self = Self(0x7fff);

    /// Construct a mask when every set bit is documented by RTF 1.9.1.
    pub const fn from_bits(bits: u16) -> Option<Self> {
        if bits & !Self::ALL.0 == 0 {
            Some(Self(bits))
        } else {
            None
        }
    }

    /// Return the raw documented event bits.
    pub const fn bits(self) -> u16 {
        self.0
    }

    /// Return whether this mask contains every bit in `other`.
    pub const fn contains(self, other: Self) -> bool {
        self.0 & other.0 == other.0
    }
}

/// Passive printing, cleanup, and event-mask document properties.
///
/// These values are retained for round-tripping only. This crate does not
/// print with QuickDraw, clean numbering definitions, instantiate VBA
/// projects, or execute document events.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentProcessingSettings {
    /// Whether `fracwidth` was present.
    pub fractional_character_widths_for_printing: bool,
    /// Explicit `ilfomacatclnupN`, preserving omission separately from `0`.
    pub abstract_numbering_cleanup: Option<AbstractNumberingCleanupStatus>,
    /// Explicit `grfdoceventsN`, preserving omission separately from zero.
    pub event_mask: Option<DocumentEventMask>,
}

impl DocumentProcessingSettings {
    /// Return whether all three properties were omitted.
    pub fn is_empty(&self) -> bool {
        !self.fractional_character_widths_for_printing
            && self.abstract_numbering_cleanup.is_none()
            && self.event_mask.is_none()
    }

    /// Return the explicit cleanup status or the RTF omission default.
    pub fn effective_abstract_numbering_cleanup(&self) -> AbstractNumberingCleanupStatus {
        self.abstract_numbering_cleanup
            .unwrap_or(AbstractNumberingCleanupStatus::Reviewed)
    }
}
