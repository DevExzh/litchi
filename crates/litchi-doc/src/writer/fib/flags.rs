//! Typed flag words used by the writer-side File Information Block.

/// FIB base version for Word 97 compatibility.
///
/// The actual format version is stored in the new FIB short section.
pub(crate) const FIB_BASE_VERSION: u16 = 0x00C1;

/// Product version emitted by the writer.
pub(crate) const PRODUCT_VERSION: u16 = 0x0000;

/// The flag state represented by FibBase.flags1.
#[derive(Debug, Clone, Copy)]
pub(super) struct BaseFlags {
    pub(super) complex: bool,
    pub(super) glossary: bool,
    pub(super) template: bool,
}

impl Default for BaseFlags {
    fn default() -> Self {
        Self {
            complex: true,
            glossary: false,
            template: false,
        }
    }
}

impl BaseFlags {
    /// Encode the typed flags using the bit assignments from [MS-DOC].
    pub(super) fn encode(self) -> u16 {
        let mut value = 0;

        // The writer always uses the 1Table stream and extended characters.
        value |= 0x0200; // fWhichTblStm
        value |= 0x1000; // fExtChar
        if self.complex {
            value |= 0x0004; // fComplex
        }
        // Word requires cQuickSaves = 0xF for this FIB generation level.
        value |= 0x00F0;
        if self.template {
            value |= 0x0001; // fDot
        }
        if self.glossary {
            value |= 0x0002; // fGlsy
        }

        value
    }
}
