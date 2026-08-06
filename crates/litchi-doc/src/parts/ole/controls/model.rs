//! Typed OLE-control metadata owned by the `parts::ole::controls` context.

/// The OLE controls recorded in a document.
///
/// The records are inert metadata. A `Controls` value never instantiates or
/// activates an OLE control and never executes control code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Controls {
    pub(crate) controls: Vec<Control>,
    /// The on-disk stride of one `OcxInfo` entry, including `dwCookie`.
    pub(crate) entry_stride: usize,
    /// Undefined bytes following each cookie, kept in entry order.
    pub(crate) entry_padding: Vec<u8>,
}

/// One `OcxInfo` entry (MS-DOC 2.9.161).
///
/// MS-DOC defines the cookie as a unique index within the document's
/// `RgxOcxInfo` table. The specified body is decoded when present; any bytes
/// beyond that body remain opaque and are retained for round trips.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Control {
    /// Unique index of this control within the document's `RgxOcxInfo`.
    pub cookie: u32,
    /// The typed body of the specified `OcxInfo`, when the producer emitted
    /// the complete structure. Short compatibility entries remain lossless
    /// and expose `None` instead of invented field values.
    pub metadata: Option<Metadata>,
}

/// The typed body of an `OcxInfo` entry (MS-DOC 2.9.161).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Metadata {
    /// Index into the applicable field PLC.
    pub field_index: u32,
    /// Undefined accelerator-table handle retained for round trips.
    pub accelerator_handle: u32,
    /// Number of accelerator entries associated with the control.
    pub accelerator_count: u16,
    /// Control behavior flags.
    pub flags: Flags,
    /// Document substream containing the field reference.
    pub document: Document,
}

/// The defined behavior bits in an `OcxInfo` flags byte.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Flags {
    pub eats_return: bool,
    pub eats_escape: bool,
    pub default_button: bool,
    pub cancel_button: bool,
    pub failed_load: bool,
    pub right_to_left: bool,
    pub corrupt: bool,
}

impl Flags {
    pub(crate) const fn from_raw(raw: u8) -> Self {
        Self {
            eats_return: raw & (1 << 1) != 0,
            eats_escape: raw & (1 << 2) != 0,
            default_button: raw & (1 << 3) != 0,
            cancel_button: raw & (1 << 4) != 0,
            failed_load: raw & (1 << 5) != 0,
            right_to_left: raw & (1 << 6) != 0,
            corrupt: raw & (1 << 7) != 0,
        }
    }
}

/// The `idoc` document substream selector from `OcxInfo`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Document {
    Main,
    Header,
    Footnote,
    Textbox,
    Endnote,
    Comment,
    HeaderTextbox,
    Unknown(u16),
}

impl Document {
    pub(crate) const fn from_raw(raw: u16) -> Self {
        match raw {
            1 => Self::Main,
            2 => Self::Header,
            3 => Self::Footnote,
            4 => Self::Textbox,
            6 => Self::Endnote,
            7 => Self::Comment,
            8 => Self::HeaderTextbox,
            value => Self::Unknown(value),
        }
    }

    pub const fn raw(self) -> u16 {
        match self {
            Self::Main => 1,
            Self::Header => 2,
            Self::Footnote => 3,
            Self::Textbox => 4,
            Self::Endnote => 6,
            Self::Comment => 7,
            Self::HeaderTextbox => 8,
            Self::Unknown(value) => value,
        }
    }
}

impl Controls {
    pub(crate) fn from_controls(controls: Vec<Control>) -> Self {
        Self {
            controls,
            entry_stride: 4,
            entry_padding: Vec::new(),
        }
    }

    pub(crate) fn from_parts(
        controls: Vec<Control>,
        entry_stride: usize,
        entry_padding: Vec<u8>,
    ) -> Self {
        Self {
            controls,
            entry_stride,
            entry_padding,
        }
    }

    /// All recorded OLE controls, in table order.
    pub fn controls(&self) -> &[Control] {
        &self.controls
    }

    /// Number of recorded OLE controls.
    pub fn len(&self) -> usize {
        self.controls.len()
    }

    /// Whether the table contains no OLE controls.
    pub fn is_empty(&self) -> bool {
        self.controls.is_empty()
    }

    /// Number of bytes occupied by one `OcxInfo` entry in the source table.
    pub fn entry_stride(&self) -> usize {
        self.entry_stride
    }

    /// Undefined bytes following the cookie for one entry.
    ///
    /// The returned slice is empty for the four-byte canonical form. The
    /// bytes are retained so a metadata-only round trip does not normalize a
    /// Word-produced padded table.
    pub fn entry_padding(&self, index: usize) -> Option<&[u8]> {
        let padding_len = self.entry_stride.checked_sub(4)?;
        let start = index.checked_mul(padding_len)?;
        let end = start.checked_add(padding_len)?;
        self.entry_padding.get(start..end)
    }

    /// Serialize the complete `RgxOcxInfo` table without discarding padded
    /// entry bytes.
    pub fn to_bytes(&self) -> Vec<u8> {
        let padding_len = self.entry_stride.saturating_sub(4);
        let mut bytes = Vec::with_capacity(
            4usize.saturating_add(self.controls.len().saturating_mul(self.entry_stride)),
        );
        bytes.extend_from_slice(&(self.controls.len() as u32).to_le_bytes());
        for (index, control) in self.controls.iter().enumerate() {
            bytes.extend_from_slice(&control.cookie.to_le_bytes());
            if padding_len != 0 {
                let start = index.saturating_mul(padding_len);
                let end = start.saturating_add(padding_len);
                if let Some(padding) = self.entry_padding.get(start..end) {
                    bytes.extend_from_slice(padding);
                } else {
                    bytes.resize(bytes.len().saturating_add(padding_len), 0);
                }
            }
        }
        bytes
    }
}
