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

#[cfg(test)]
mod tests {
    use super::*;

    // =============================================================================
    // SrKinsokuAtom Tests
    // =============================================================================

    #[test]
    fn test_sr_kinsoku_atom_default() {
        let atom = SrKinsokuAtom::DEFAULT;
        assert_eq!(atom.kinsoku_type, 1);
    }

    #[test]
    fn test_sr_kinsoku_atom_custom_type() {
        let atom = SrKinsokuAtom { kinsoku_type: 3 };
        assert_eq!(atom.kinsoku_type, 3);
    }

    #[test]
    fn test_sr_kinsoku_atom_to_bytes() {
        let atom = SrKinsokuAtom { kinsoku_type: 1 };
        let bytes = atom.to_bytes();
        assert_eq!(bytes.len(), 4);
        assert_eq!(u32::from_le_bytes(bytes), 1);
    }

    #[test]
    fn test_sr_kinsoku_atom_all_types() {
        // Test all CJK kinsoku types (Japanese, Korean, Simplified Chinese, Traditional Chinese)
        for kinsoku_type in [1u32, 2, 3, 4] {
            let atom = SrKinsokuAtom { kinsoku_type };
            let bytes = atom.to_bytes();
            assert_eq!(u32::from_le_bytes(bytes), kinsoku_type);
        }
    }

    // =============================================================================
    // TxCFStyleAtom Tests
    // =============================================================================

    #[test]
    fn test_tx_cf_style_atom_default() {
        let atom = TxCFStyleAtom::DEFAULT;
        assert_eq!(atom.cf_mask, 0x0080);
        assert_eq!(atom.cf_flags, 0x0040);
        assert_eq!(atom.reserved, 0x0000);
        assert_eq!(atom.font_ref, 0xFFFF);
    }

    #[test]
    fn test_tx_cf_style_atom_to_bytes() {
        let atom = TxCFStyleAtom::DEFAULT;
        let bytes = atom.to_bytes();
        assert_eq!(bytes.len(), 8);

        // Verify byte layout
        assert_eq!(u16::from_le_bytes([bytes[0], bytes[1]]), 0x0080);
        assert_eq!(u16::from_le_bytes([bytes[2], bytes[3]]), 0x0040);
        assert_eq!(u16::from_le_bytes([bytes[4], bytes[5]]), 0x0000);
        assert_eq!(u16::from_le_bytes([bytes[6], bytes[7]]), 0xFFFF);
    }

    #[test]
    fn test_tx_cf_style_atom_custom_values() {
        let atom = TxCFStyleAtom {
            cf_mask: 0x00FF,
            cf_flags: 0x00AA,
            reserved: 0x1234,
            font_ref: 0x0005,
        };
        let bytes = atom.to_bytes();
        assert_eq!(u16::from_le_bytes([bytes[0], bytes[1]]), 0x00FF);
        assert_eq!(u16::from_le_bytes([bytes[2], bytes[3]]), 0x00AA);
        assert_eq!(u16::from_le_bytes([bytes[4], bytes[5]]), 0x1234);
        assert_eq!(u16::from_le_bytes([bytes[6], bytes[7]]), 0x0005);
    }

    // =============================================================================
    // TxPFStyleAtom Tests
    // =============================================================================

    #[test]
    fn test_tx_pf_style_atom_default() {
        let atom = TxPFStyleAtom::DEFAULT;
        assert_eq!(atom.pf_mask, 0x0800_0000);
        assert_eq!(atom.bullet_char, 0x2E);
        assert_eq!(atom.pf_flags, 0x0000_0002);
    }

    #[test]
    fn test_tx_pf_style_atom_to_bytes() {
        let atom = TxPFStyleAtom::DEFAULT;
        let bytes = atom.to_bytes();
        assert_eq!(bytes.len(), 12);

        assert_eq!(
            u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]),
            0x0800_0000
        );
        assert_eq!(
            u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]),
            0x2E
        );
        assert_eq!(
            u32::from_le_bytes([bytes[8], bytes[9], bytes[10], bytes[11]]),
            0x0000_0002
        );
    }

    #[test]
    fn test_tx_pf_style_atom_custom_bullet() {
        let atom = TxPFStyleAtom {
            pf_mask: 0x1234_5678,
            bullet_char: 0x2022, // Unicode bullet
            pf_flags: 0x8765_4321,
        };
        let bytes = atom.to_bytes();
        assert_eq!(
            u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]),
            0x2022
        );
    }

    // =============================================================================
    // TxSIStyleAtom Tests
    // =============================================================================

    #[test]
    fn test_tx_si_style_atom_default() {
        let atom = TxSIStyleAtom::DEFAULT;
        assert_eq!(atom.si_mask, 0x0000_0007);
        assert_eq!(atom.lang, lang_id::NEUTRAL);
        assert_eq!(atom.alt_lang, lang_id::EN_US);
        assert_eq!(atom.reserved, 0);
    }

    #[test]
    fn test_tx_si_style_atom_to_bytes() {
        let atom = TxSIStyleAtom::DEFAULT;
        let bytes = atom.to_bytes();
        assert_eq!(bytes.len(), 10);

        assert_eq!(
            u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]),
            0x0000_0007
        );
        assert_eq!(u16::from_le_bytes([bytes[4], bytes[5]]), lang_id::NEUTRAL);
        assert_eq!(u16::from_le_bytes([bytes[6], bytes[7]]), lang_id::EN_US);
        assert_eq!(u16::from_le_bytes([bytes[8], bytes[9]]), 0);
    }

    #[test]
    fn test_tx_si_style_atom_custom_lang() {
        let atom = TxSIStyleAtom {
            si_mask: 0x0000_00FF,
            lang: 0x0407,     // German
            alt_lang: 0x040C, // French
            reserved: 0x1234,
        };
        let bytes = atom.to_bytes();
        assert_eq!(u16::from_le_bytes([bytes[4], bytes[5]]), 0x0407);
        assert_eq!(u16::from_le_bytes([bytes[6], bytes[7]]), 0x040C);
        assert_eq!(u16::from_le_bytes([bytes[8], bytes[9]]), 0x1234);
    }

    #[test]
    fn test_lang_id_constants() {
        assert_eq!(lang_id::EN_US, 0x0409);
        assert_eq!(lang_id::NEUTRAL, 0x0002);
    }

    // =============================================================================
    // SheetPropertiesAtom Tests
    // =============================================================================

    #[test]
    fn test_sheet_properties_atom_default() {
        let atom = SheetPropertiesAtom::DEFAULT;
        assert_eq!(atom.creation_time, 0x3B9A_CA00_F6B0_93BA);
        assert_eq!(atom.modification_time, 0x3B9A_CA00_C794_07AD);
        assert_eq!(atom.flags, 0x0101);
        assert_eq!(atom.reserved, 0x0000);
    }

    #[test]
    fn test_sheet_properties_atom_to_bytes() {
        let atom = SheetPropertiesAtom::DEFAULT;
        let bytes = atom.to_bytes();
        assert_eq!(bytes.len(), 20);

        assert_eq!(
            u64::from_le_bytes([
                bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7]
            ]),
            0x3B9A_CA00_F6B0_93BA
        );
        assert_eq!(
            u64::from_le_bytes([
                bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14],
                bytes[15]
            ]),
            0x3B9A_CA00_C794_07AD
        );
        assert_eq!(u16::from_le_bytes([bytes[16], bytes[17]]), 0x0101);
        assert_eq!(u16::from_le_bytes([bytes[18], bytes[19]]), 0x0000);
    }

    #[test]
    fn test_sheet_properties_atom_custom_timestamps() {
        let atom = SheetPropertiesAtom {
            creation_time: 0x0000_0000_0000_0000,
            modification_time: 0xFFFF_FFFF_FFFF_FFFF,
            flags: 0x1234,
            reserved: 0x5678,
        };
        let bytes = atom.to_bytes();
        assert_eq!(
            u64::from_le_bytes([
                bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7]
            ]),
            0
        );
        assert_eq!(
            u64::from_le_bytes([
                bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14],
                bytes[15]
            ]),
            u64::MAX
        );
    }

    #[test]
    fn test_sheet_properties_child_type_constant() {
        assert_eq!(SHEET_PROPERTIES_CHILD_TYPE, 0x0415);
    }

    // =============================================================================
    // SlideViewInfoAtom Tests
    // =============================================================================

    #[test]
    fn test_slide_view_info_atom_default() {
        let atom = SlideViewInfoAtom::DEFAULT;
        assert!(!atom.snap_to_grid);
        assert!(atom.snap_to_shape);
        assert!(!atom.show_guides);
    }

    #[test]
    fn test_slide_view_info_atom_to_bytes() {
        let atom = SlideViewInfoAtom::DEFAULT;
        let bytes = atom.to_bytes();
        assert_eq!(bytes.len(), 3);
        assert_eq!(bytes[0], 0); // snap_to_grid = false
        assert_eq!(bytes[1], 1); // snap_to_shape = true
        assert_eq!(bytes[2], 0); // show_guides = false
    }

    #[test]
    fn test_slide_view_info_atom_all_enabled() {
        let atom = SlideViewInfoAtom {
            snap_to_grid: true,
            snap_to_shape: true,
            show_guides: true,
        };
        let bytes = atom.to_bytes();
        assert_eq!(bytes[0], 1);
        assert_eq!(bytes[1], 1);
        assert_eq!(bytes[2], 1);
    }

    #[test]
    fn test_slide_view_info_atom_all_disabled() {
        let atom = SlideViewInfoAtom {
            snap_to_grid: false,
            snap_to_shape: false,
            show_guides: false,
        };
        let bytes = atom.to_bytes();
        assert_eq!(bytes[0], 0);
        assert_eq!(bytes[1], 0);
        assert_eq!(bytes[2], 0);
    }

    // =============================================================================
    // VBAInfoAtom Tests
    // =============================================================================

    #[test]
    fn test_vba_info_atom_default() {
        let atom = VBAInfoAtom::DEFAULT;
        assert_eq!(atom.persist_id_ref, 0);
        assert_eq!(atom.flags, vba_flags::HAS_PROJECT);
    }

    #[test]
    fn test_vba_info_atom_to_bytes() {
        let atom = VBAInfoAtom::DEFAULT;
        let bytes = atom.to_bytes();
        assert_eq!(bytes.len(), 12);

        assert_eq!(
            u64::from_le_bytes([
                bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7]
            ]),
            0
        );
        assert_eq!(
            u32::from_le_bytes([bytes[8], bytes[9], bytes[10], bytes[11]]),
            vba_flags::HAS_PROJECT
        );
    }

    #[test]
    fn test_vba_info_atom_with_macros() {
        let atom = VBAInfoAtom {
            persist_id_ref: 12345,
            flags: vba_flags::HAS_MACROS | vba_flags::HAS_PROJECT,
        };
        let bytes = atom.to_bytes();
        assert_eq!(
            u64::from_le_bytes([
                bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7]
            ]),
            12345
        );
        assert_eq!(
            u32::from_le_bytes([bytes[8], bytes[9], bytes[10], bytes[11]]),
            0x0000_0003
        );
    }

    #[test]
    fn test_vba_flags_constants() {
        assert_eq!(vba_flags::HAS_MACROS, 0x0000_0001);
        assert_eq!(vba_flags::HAS_PROJECT, 0x0000_0002);
    }

    // =============================================================================
    // Clone and Debug Tests
    // =============================================================================

    #[test]
    fn test_sr_kinsoku_atom_clone() {
        let atom = SrKinsokuAtom { kinsoku_type: 2 };
        let cloned = atom.clone();
        assert_eq!(atom.kinsoku_type, cloned.kinsoku_type);
    }

    #[test]
    fn test_tx_cf_style_atom_clone() {
        let atom = TxCFStyleAtom::DEFAULT;
        let cloned = atom.clone();
        assert_eq!(atom.cf_mask, cloned.cf_mask);
        assert_eq!(atom.font_ref, cloned.font_ref);
    }

    #[test]
    fn test_tx_pf_style_atom_clone() {
        let atom = TxPFStyleAtom::DEFAULT;
        let cloned = atom.clone();
        assert_eq!(atom.pf_mask, cloned.pf_mask);
        assert_eq!(atom.bullet_char, cloned.bullet_char);
    }

    #[test]
    fn test_tx_si_style_atom_clone() {
        let atom = TxSIStyleAtom::DEFAULT;
        let cloned = atom.clone();
        assert_eq!(atom.lang, cloned.lang);
        assert_eq!(atom.alt_lang, cloned.alt_lang);
    }

    #[test]
    fn test_sheet_properties_atom_clone() {
        let atom = SheetPropertiesAtom::DEFAULT;
        let cloned = atom.clone();
        assert_eq!(atom.creation_time, cloned.creation_time);
        assert_eq!(atom.flags, cloned.flags);
    }

    #[test]
    fn test_slide_view_info_atom_clone() {
        let atom = SlideViewInfoAtom::DEFAULT;
        let cloned = atom.clone();
        assert_eq!(atom.snap_to_grid, cloned.snap_to_grid);
        assert_eq!(atom.snap_to_shape, cloned.snap_to_shape);
    }

    #[test]
    fn test_vba_info_atom_clone() {
        let atom = VBAInfoAtom::DEFAULT;
        let cloned = atom.clone();
        assert_eq!(atom.persist_id_ref, cloned.persist_id_ref);
        assert_eq!(atom.flags, cloned.flags);
    }

    #[test]
    fn test_debug_formatting() {
        let sr = SrKinsokuAtom::DEFAULT;
        let cf = TxCFStyleAtom::DEFAULT;
        let pf = TxPFStyleAtom::DEFAULT;
        let si = TxSIStyleAtom::DEFAULT;
        let sheet = SheetPropertiesAtom::DEFAULT;
        let view = SlideViewInfoAtom::DEFAULT;
        let vba = VBAInfoAtom::DEFAULT;

        assert!(format!("{:?}", sr).contains("SrKinsokuAtom"));
        assert!(format!("{:?}", cf).contains("TxCFStyleAtom"));
        assert!(format!("{:?}", pf).contains("TxPFStyleAtom"));
        assert!(format!("{:?}", si).contains("TxSIStyleAtom"));
        assert!(format!("{:?}", sheet).contains("SheetPropertiesAtom"));
        assert!(format!("{:?}", view).contains("SlideViewInfoAtom"));
        assert!(format!("{:?}", vba).contains("VBAInfoAtom"));
    }
}
