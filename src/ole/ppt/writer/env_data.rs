//! Environment container data structures (MS-PPT 2.4.4)
//!
//! Structured types for Environment child atoms based on MS-PPT specification.

// =============================================================================
// SrKinsokuAtom (MS-PPT 2.9.26)
// =============================================================================

/// SrKinsokuAtom - Line breaking rules for CJK text
#[derive(Debug, Clone, Copy)]
pub struct SrKinsokuAtom {
    /// Kinsoku type: 1 = Japanese, 2 = Korean, 3 = Simplified Chinese, 4 = Traditional Chinese
    pub kinsoku_type: u32,
}

impl SrKinsokuAtom {
    /// Default: Japanese line breaking rules
    pub const DEFAULT: Self = Self { kinsoku_type: 1 };

    pub fn to_bytes(&self) -> [u8; 4] {
        self.kinsoku_type.to_le_bytes()
    }
}

// =============================================================================
// TxCFStyleAtom (MS-PPT 2.9.52) - Character Formatting Defaults
// =============================================================================

/// TxCFStyleAtom - Default character formatting for text
#[derive(Debug, Clone, Copy)]
pub struct TxCFStyleAtom {
    /// Mask indicating which fields are valid (MS-PPT 2.9.6 TextCFException)
    pub cf_mask: u16,
    /// Character formatting flags
    pub cf_flags: u16,
    /// Reserved field
    pub reserved: u16,
    /// Font reference index (0xFFFF = no font specified)
    pub font_ref: u16,
}

impl TxCFStyleAtom {
    /// Default character formatting from POI
    pub const DEFAULT: Self = Self {
        cf_mask: 0x0080,  // fontRef field is valid
        cf_flags: 0x0040, // formatting flags
        reserved: 0x0000,
        font_ref: 0xFFFF, // no font specified
    };

    pub fn to_bytes(&self) -> [u8; 8] {
        let mut data = [0u8; 8];
        data[0..2].copy_from_slice(&self.cf_mask.to_le_bytes());
        data[2..4].copy_from_slice(&self.cf_flags.to_le_bytes());
        data[4..6].copy_from_slice(&self.reserved.to_le_bytes());
        data[6..8].copy_from_slice(&self.font_ref.to_le_bytes());
        data
    }
}

// =============================================================================
// TxPFStyleAtom (MS-PPT 2.9.53) - Paragraph Formatting Defaults
// =============================================================================

/// TxPFStyleAtom - Default paragraph formatting for text
#[derive(Debug, Clone, Copy)]
pub struct TxPFStyleAtom {
    /// Mask indicating which fields are valid (MS-PPT 2.9.18 TextPFException)
    pub pf_mask: u32,
    /// Bullet character (Unicode code point)
    pub bullet_char: u32,
    /// Paragraph formatting flags
    pub pf_flags: u32,
}

impl TxPFStyleAtom {
    /// Default paragraph formatting from POI
    pub const DEFAULT: Self = Self {
        pf_mask: 0x0800_0000,  // bulletChar field is valid
        bullet_char: 0x2E,     // '.' character
        pf_flags: 0x0000_0002, // paragraph flags
    };

    pub fn to_bytes(&self) -> [u8; 12] {
        let mut data = [0u8; 12];
        data[0..4].copy_from_slice(&self.pf_mask.to_le_bytes());
        data[4..8].copy_from_slice(&self.bullet_char.to_le_bytes());
        data[8..12].copy_from_slice(&self.pf_flags.to_le_bytes());
        data
    }
}

// =============================================================================
// TxSIStyleAtom (MS-PPT 2.9.54) - Special Info Formatting
// =============================================================================

/// Language ID constants (MS-LCID)
pub mod lang_id {
    /// English (United States)
    pub const EN_US: u16 = 0x0409;
    /// Neutral/Default
    pub const NEUTRAL: u16 = 0x0002;
}

/// TxSIStyleAtom - Special text info (language, spell check)
#[derive(Debug, Clone, Copy)]
pub struct TxSIStyleAtom {
    /// Mask indicating which fields are valid
    pub si_mask: u32,
    /// Primary language ID
    pub lang: u16,
    /// Alternate language ID (for spell checking)
    pub alt_lang: u16,
    /// Reserved
    pub reserved: u16,
}

impl TxSIStyleAtom {
    /// Default special info from POI
    pub const DEFAULT: Self = Self {
        si_mask: 0x0000_0007, // lang and altLang fields valid
        lang: lang_id::NEUTRAL,
        alt_lang: lang_id::EN_US,
        reserved: 0,
    };

    pub fn to_bytes(&self) -> [u8; 10] {
        let mut data = [0u8; 10];
        data[0..4].copy_from_slice(&self.si_mask.to_le_bytes());
        data[4..6].copy_from_slice(&self.lang.to_le_bytes());
        data[6..8].copy_from_slice(&self.alt_lang.to_le_bytes());
        data[8..10].copy_from_slice(&self.reserved.to_le_bytes());
        data
    }
}

// =============================================================================
// SheetPropertiesAtom (undocumented, reverse-engineered from POI)
// =============================================================================

/// SheetPropertiesAtom - Document timestamps and flags
#[derive(Debug, Clone, Copy)]
pub struct SheetPropertiesAtom {
    /// Creation timestamp (Windows FILETIME)
    pub creation_time: u64,
    /// Last modification timestamp (Windows FILETIME)
    pub modification_time: u64,
    /// Sheet flags (interpretation unknown)
    pub flags: u16,
    /// Reserved
    pub reserved: u16,
}

impl SheetPropertiesAtom {
    /// Default timestamps from POI empty.ppt
    pub const DEFAULT: Self = Self {
        creation_time: 0x3B9A_CA00_F6B0_93BA,
        modification_time: 0x3B9A_CA00_C794_07AD,
        flags: 0x0101,
        reserved: 0x0000,
    };

    pub fn to_bytes(&self) -> [u8; 20] {
        let mut data = [0u8; 20];
        data[0..8].copy_from_slice(&self.creation_time.to_le_bytes());
        data[8..16].copy_from_slice(&self.modification_time.to_le_bytes());
        data[16..18].copy_from_slice(&self.flags.to_le_bytes());
        data[18..20].copy_from_slice(&self.reserved.to_le_bytes());
        data
    }
}

/// SheetProperties child atom type (undocumented)
pub const SHEET_PROPERTIES_CHILD_TYPE: u16 = 0x0415;

// =============================================================================
// SlideViewInfoAtom (MS-PPT 2.4.21.3)
// =============================================================================

/// SlideViewInfoAtom - View state for slide editing
#[derive(Debug, Clone, Copy)]
pub struct SlideViewInfoAtom {
    /// Snap to grid enabled
    pub snap_to_grid: bool,
    /// Snap to shapes enabled  
    pub snap_to_shape: bool,
    /// Show guides
    pub show_guides: bool,
}

impl SlideViewInfoAtom {
    /// Default view settings from POI
    pub const DEFAULT: Self = Self {
        snap_to_grid: false,
        snap_to_shape: true,
        show_guides: false,
    };

    pub fn to_bytes(&self) -> [u8; 3] {
        [
            self.snap_to_grid as u8,
            self.snap_to_shape as u8,
            self.show_guides as u8,
        ]
    }
}

// =============================================================================
// VBAInfoAtom (MS-PPT 2.10.1)
// =============================================================================

/// VBA macro flags
pub mod vba_flags {
    /// Document has macros (fHasMacros)
    pub const HAS_MACROS: u32 = 0x0000_0001;
    /// Macros are enabled (fHasProject)
    pub const HAS_PROJECT: u32 = 0x0000_0002;
}

/// VBAInfoAtom - VBA macro information
#[derive(Debug, Clone, Copy)]
pub struct VBAInfoAtom {
    /// Persist ID reference to VBA storage (0 if none)
    pub persist_id_ref: u64,
    /// VBA flags (see vba_flags module)
    pub flags: u32,
}

impl VBAInfoAtom {
    /// Default: no VBA but project flag set (POI behavior)
    pub const DEFAULT: Self = Self {
        persist_id_ref: 0,
        flags: vba_flags::HAS_PROJECT,
    };

    pub fn to_bytes(&self) -> [u8; 12] {
        let mut data = [0u8; 12];
        data[0..8].copy_from_slice(&self.persist_id_ref.to_le_bytes());
        data[8..12].copy_from_slice(&self.flags.to_le_bytes());
        data
    }
}
