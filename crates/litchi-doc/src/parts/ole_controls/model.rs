//! Typed, lossless `OcxInfo` metadata.

use crate::package::{Error as PackageError, Result};
use std::collections::HashSet;

/// The fixed serialized size of one `OcxInfo` record (MS-DOC 2.9.161).
pub(super) const OCX_INFO_SIZE: usize = 20;

/// The location of an OLE-control field in the document stories (the
/// `OcxInfo.idoc` value from MS-DOC 2.9.161).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum Story {
    /// Main document story (`idoc = 1`).
    Main = 1,
    /// Header story (`idoc = 2`).
    Header = 2,
    /// Footnote story (`idoc = 3`).
    Footnote = 3,
    /// Textbox story (`idoc = 4`).
    Textbox = 4,
    /// Endnote story (`idoc = 6`).
    Endnote = 6,
    /// Comment story (`idoc = 7`).
    Comment = 7,
    /// Header-textbox story (`idoc = 8`).
    HeaderTextbox = 8,
}

impl Story {
    /// Decode the on-disk `idoc` value.
    pub(crate) fn from_raw(value: u16) -> Result<Self> {
        match value {
            1 => Ok(Self::Main),
            2 => Ok(Self::Header),
            3 => Ok(Self::Footnote),
            4 => Ok(Self::Textbox),
            6 => Ok(Self::Endnote),
            7 => Ok(Self::Comment),
            8 => Ok(Self::HeaderTextbox),
            _ => Err(corrupted(format!("OcxInfo idoc value {value} is invalid"))),
        }
    }

    /// The exact on-disk `idoc` value.
    pub const fn raw(self) -> u16 {
        self as u16
    }
}

/// The `OcxInfo` flag word.
///
/// The unknown high byte is retained as `reserved_bits`; `hAccel` and the
/// second reserved word are retained by [`OcxInfo`] as well. MS-DOC marks
/// these values as undefined/ignored rather than requiring them to be zero,
/// so parsing does not discard or normalize them.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Flags {
    raw: u16,
}

impl Flags {
    const FIFLD: u16 = 1 << 0;

    /// Construct the flag word from its individual semantic bits.
    #[allow(clippy::too_many_arguments)]
    pub const fn new(
        eats_return: bool,
        eats_escape: bool,
        default_button: bool,
        cancel_button: bool,
        failed_load: bool,
        right_to_left: bool,
        corrupt: bool,
        reserved_bits: u8,
    ) -> Self {
        let mut raw = Self::FIFLD;
        if eats_return {
            raw |= 1 << 1;
        }
        if eats_escape {
            raw |= 1 << 2;
        }
        if default_button {
            raw |= 1 << 3;
        }
        if cancel_button {
            raw |= 1 << 4;
        }
        if failed_load {
            raw |= 1 << 5;
        }
        if right_to_left {
            raw |= 1 << 6;
        }
        if corrupt {
            raw |= 1 << 7;
        }
        Self {
            raw: raw | (reserved_bits as u16) << 8,
        }
    }

    /// Decode the raw flag word, enforcing the required `fifld` bit.
    pub(crate) fn from_raw(raw: u16) -> Result<Self> {
        if raw & Self::FIFLD == 0 {
            return Err(corrupted("OcxInfo fifld must be set"));
        }
        Ok(Self { raw })
    }

    /// The exact serialized flag word.
    pub const fn raw(self) -> u16 {
        self.raw
    }

    /// Whether the record is associated with a field (`fifld`). Always true
    /// for a valid value because MS-DOC requires this bit to be set.
    pub const fn field_present(self) -> bool {
        self.raw & Self::FIFLD != 0
    }

    /// Whether the control consumes ENTER.
    pub const fn eats_return(self) -> bool {
        self.raw & (1 << 1) != 0
    }

    /// Whether the control consumes ESC.
    pub const fn eats_escape(self) -> bool {
        self.raw & (1 << 2) != 0
    }

    /// Whether the control is the default button.
    pub const fn default_button(self) -> bool {
        self.raw & (1 << 3) != 0
    }

    /// Whether the control is the default CANCEL button.
    pub const fn cancel_button(self) -> bool {
        self.raw & (1 << 4) != 0
    }

    /// Whether loading the control failed.
    pub const fn failed_load(self) -> bool {
        self.raw & (1 << 5) != 0
    }

    /// Whether the control uses right-to-left display handling.
    pub const fn right_to_left(self) -> bool {
        self.raw & (1 << 6) != 0
    }

    /// Whether the control is marked corrupt.
    pub const fn corrupt(self) -> bool {
        self.raw & (1 << 7) != 0
    }

    /// The ignored high-byte bits, retained losslessly.
    pub const fn reserved_bits(self) -> u8 {
        (self.raw >> 8) as u8
    }
}

/// One fixed-size `OcxInfo` record (MS-DOC 2.9.161).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct OcxInfo {
    cookie: u32,
    field_index: u32,
    accelerator_handle: u32,
    accelerator_count: u16,
    flags: Flags,
    story: Story,
    reserved: u16,
}

impl OcxInfo {
    /// Construct a valid record. The ignored/reserved values are retained
    /// exactly when the record is serialized.
    pub const fn new(
        cookie: u32,
        field_index: u32,
        accelerator_handle: u32,
        accelerator_count: u16,
        flags: Flags,
        story: Story,
        reserved: u16,
    ) -> Self {
        Self {
            cookie,
            field_index,
            accelerator_handle,
            accelerator_count,
            flags,
            story,
            reserved,
        }
    }

    /// Unique `dwCookie` index in the containing table.
    pub const fn cookie(self) -> u32 {
        self.cookie
    }

    /// `ifld`, the field index in the story selected by [`Self::story`].
    pub const fn field_index(self) -> u32 {
        self.field_index
    }

    /// Undefined `hAccel`, retained without interpretation.
    pub const fn accelerator_handle(self) -> u32 {
        self.accelerator_handle
    }

    /// Number of accelerator entries (`cAccel`).
    pub const fn accelerator_count(self) -> u16 {
        self.accelerator_count
    }

    /// Semantic and retained bits from the record flag word.
    pub const fn flags(self) -> Flags {
        self.flags
    }

    /// Story containing the field referenced by `ifld`.
    pub const fn story(self) -> Story {
        self.story
    }

    /// Undefined `reserved2`, retained without interpretation.
    pub const fn reserved(self) -> u16 {
        self.reserved
    }
}

/// The `RgxOcxInfo` array (MS-DOC 2.9.229).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RgxOcxInfo {
    infos: Vec<OcxInfo>,
}

impl RgxOcxInfo {
    /// Construct an array and validate the document-wide cookie invariant.
    pub fn try_new(infos: Vec<OcxInfo>) -> Result<Self> {
        validate_cookies(&infos)?;
        u32::try_from(infos.len()).map_err(|_| corrupted("RgxOcxInfo count exceeds u32::MAX"))?;
        Ok(Self { infos })
    }

    pub(crate) fn from_infos(infos: Vec<OcxInfo>) -> Self {
        Self { infos }
    }

    /// Records in their original table order.
    pub fn infos(&self) -> &[OcxInfo] {
        &self.infos
    }

    /// Number of records.
    pub fn len(&self) -> usize {
        self.infos.len()
    }

    /// Whether no records are present.
    pub fn is_empty(&self) -> bool {
        self.infos.is_empty()
    }

    pub(crate) fn validate(&self) -> Result<()> {
        validate_cookies(&self.infos)?;
        u32::try_from(self.infos.len())
            .map_err(|_| corrupted("RgxOcxInfo count exceeds u32::MAX"))?;
        Ok(())
    }
}

pub(crate) fn validate_cookies(infos: &[OcxInfo]) -> Result<()> {
    let mut cookies = HashSet::with_capacity(infos.len());
    for info in infos {
        if !cookies.insert(info.cookie()) {
            return Err(corrupted("OcxInfo dwCookie values must be unique"));
        }
    }
    Ok(())
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
