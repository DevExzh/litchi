//! Constants from [MS-WMF].
//!
//! WMF function values are not interchangeable with Win32 API constants.  In
//! particular, the high byte of a few variable-sized record functions can be
//! rewritten by producers.  [`record::canonical`] deals with those records.

/// WMF record function codes.
#[allow(dead_code)]
pub mod record {
    pub const EOF: u16 = 0x0000;
    pub const SAVE_DC: u16 = 0x001e;
    pub const REALIZE_PALETTE: u16 = 0x0035;
    pub const SET_PALETTE_ENTRIES: u16 = 0x0037;
    pub const CREATE_PALETTE: u16 = 0x00f7;
    pub const SET_BK_MODE: u16 = 0x0102;
    pub const SET_MAP_MODE: u16 = 0x0103;
    pub const SET_ROP2: u16 = 0x0104;
    pub const SET_REL_ABS: u16 = 0x0105;
    pub const SET_POLY_FILL_MODE: u16 = 0x0106;
    pub const SET_STRETCH_BLT_MODE: u16 = 0x0107;
    pub const SET_TEXT_CHAR_EXTRA: u16 = 0x0108;
    pub const RESTORE_DC: u16 = 0x0127;
    pub const INVERT_REGION: u16 = 0x012a;
    pub const PAINT_REGION: u16 = 0x012b;
    pub const SELECT_CLIP_REGION: u16 = 0x012c;
    pub const SELECT_OBJECT: u16 = 0x012d;
    pub const SET_TEXT_ALIGN: u16 = 0x012e;
    pub const RESIZE_PALETTE: u16 = 0x0139;
    pub const CREATE_DIB_PATTERN_BRUSH: u16 = 0x0142;
    pub const DIB_CREATE_PATTERN_BRUSH: u16 = CREATE_DIB_PATTERN_BRUSH;
    pub const SET_LAYOUT: u16 = 0x0149;
    pub const DELETE_OBJECT: u16 = 0x01f0;
    pub const CREATE_PATTERN_BRUSH: u16 = 0x01f9;
    pub const SET_BK_COLOR: u16 = 0x0201;
    pub const SET_TEXT_COLOR: u16 = 0x0209;
    pub const SET_TEXT_JUSTIFICATION: u16 = 0x020a;
    pub const SET_WINDOW_ORG: u16 = 0x020b;
    pub const SET_WINDOW_EXT: u16 = 0x020c;
    pub const SET_VIEWPORT_ORG: u16 = 0x020d;
    pub const SET_VIEWPORT_EXT: u16 = 0x020e;
    pub const OFFSET_WINDOW_ORG: u16 = 0x020f;
    pub const OFFSET_VIEWPORT_ORG: u16 = 0x0211;
    pub const LINE_TO: u16 = 0x0213;
    pub const MOVE_TO: u16 = 0x0214;
    pub const OFFSET_CLIP_RGN: u16 = 0x0220;
    pub const FILL_REGION: u16 = 0x0228;
    pub const SET_MAPPER_FLAGS: u16 = 0x0231;
    pub const SELECT_PALETTE: u16 = 0x0234;
    pub const CREATE_PEN_INDIRECT: u16 = 0x02fa;
    pub const CREATE_FONT_INDIRECT: u16 = 0x02fb;
    pub const CREATE_BRUSH_INDIRECT: u16 = 0x02fc;
    pub const POLYGON: u16 = 0x0324;
    pub const POLYLINE: u16 = 0x0325;
    pub const SCALE_WINDOW_EXT: u16 = 0x0410;
    pub const SCALE_VIEWPORT_EXT: u16 = 0x0412;
    pub const EXCLUDE_CLIP_RECT: u16 = 0x0415;
    pub const INTERSECT_CLIP_RECT: u16 = 0x0416;
    pub const ELLIPSE: u16 = 0x0418;
    pub const FLOOD_FILL: u16 = 0x0419;
    pub const RECTANGLE: u16 = 0x041b;
    pub const SET_PIXEL: u16 = 0x041f;
    /// Historical name retained for source compatibility.
    pub const SET_PIXEL_V: u16 = SET_PIXEL;
    pub const FRAME_REGION: u16 = 0x0429;
    pub const ANIMATE_PALETTE: u16 = 0x0436;
    pub const TEXT_OUT: u16 = 0x0521;
    pub const POLYPOLYGON: u16 = 0x0538;
    pub const EXT_FLOOD_FILL: u16 = 0x0548;
    pub const ROUND_RECT: u16 = 0x061c;
    pub const PAT_BLT: u16 = 0x061d;
    pub const PATBLT: u16 = PAT_BLT;
    pub const ESCAPE: u16 = 0x0626;
    pub const CREATE_REGION: u16 = 0x06ff;
    pub const ARC: u16 = 0x0817;
    pub const PIE: u16 = 0x081a;
    pub const CHORD: u16 = 0x0830;
    pub const BIT_BLT: u16 = 0x0922;
    pub const DIB_BIT_BLT: u16 = 0x0940;
    pub const EXT_TEXT_OUT: u16 = 0x0a32;
    pub const STRETCH_BLT: u16 = 0x0b23;
    pub const DIB_STRETCH_BLT: u16 = 0x0b41;
    pub const SET_DIB_TO_DEV: u16 = 0x0d33;
    pub const STRETCH_DIB: u16 = 0x0f43;

    /// Every record type in the MS-WMF RecordType enumeration, without aliases.
    pub const ALL: &[u16] = &[
        EOF,
        REALIZE_PALETTE,
        SET_PALETTE_ENTRIES,
        SET_BK_MODE,
        SET_MAP_MODE,
        SET_ROP2,
        SET_REL_ABS,
        SET_POLY_FILL_MODE,
        SET_STRETCH_BLT_MODE,
        SET_TEXT_CHAR_EXTRA,
        RESTORE_DC,
        RESIZE_PALETTE,
        CREATE_DIB_PATTERN_BRUSH,
        SET_LAYOUT,
        SET_BK_COLOR,
        SET_TEXT_COLOR,
        OFFSET_VIEWPORT_ORG,
        LINE_TO,
        MOVE_TO,
        OFFSET_CLIP_RGN,
        FILL_REGION,
        SET_MAPPER_FLAGS,
        SELECT_PALETTE,
        POLYGON,
        POLYLINE,
        SET_TEXT_JUSTIFICATION,
        SET_WINDOW_ORG,
        SET_WINDOW_EXT,
        SET_VIEWPORT_ORG,
        SET_VIEWPORT_EXT,
        OFFSET_WINDOW_ORG,
        SCALE_WINDOW_EXT,
        SCALE_VIEWPORT_EXT,
        EXCLUDE_CLIP_RECT,
        INTERSECT_CLIP_RECT,
        ELLIPSE,
        FLOOD_FILL,
        FRAME_REGION,
        ANIMATE_PALETTE,
        TEXT_OUT,
        POLYPOLYGON,
        EXT_FLOOD_FILL,
        RECTANGLE,
        SET_PIXEL,
        ROUND_RECT,
        PAT_BLT,
        SAVE_DC,
        PIE,
        STRETCH_BLT,
        ESCAPE,
        INVERT_REGION,
        PAINT_REGION,
        SELECT_CLIP_REGION,
        SELECT_OBJECT,
        SET_TEXT_ALIGN,
        ARC,
        CHORD,
        BIT_BLT,
        EXT_TEXT_OUT,
        SET_DIB_TO_DEV,
        DIB_BIT_BLT,
        DIB_STRETCH_BLT,
        STRETCH_DIB,
        DELETE_OBJECT,
        CREATE_PALETTE,
        CREATE_PATTERN_BRUSH,
        CREATE_PEN_INDIRECT,
        CREATE_FONT_INDIRECT,
        CREATE_BRUSH_INDIRECT,
        CREATE_REGION,
    ];

    /// Normalize record functions whose high byte can encode a parameter
    /// count.  Other records retain their complete function value.
    #[inline]
    pub const fn canonical(function: u16) -> u16 {
        match function & 0x00ff {
            0x22 => BIT_BLT,
            0x23 => STRETCH_BLT,
            0x24 => POLYGON,
            0x25 => POLYLINE,
            0x37 => SET_PALETTE_ENTRIES,
            0x40 => DIB_BIT_BLT,
            0x41 => DIB_STRETCH_BLT,
            _ => function,
        }
    }
}

/// PenStyle flags.
#[allow(dead_code)]
pub mod pen {
    pub const PS_STYLE_MASK: u16 = 0x000f;
    pub const PS_SOLID: u16 = 0;
    pub const PS_DASH: u16 = 1;
    pub const PS_DOT: u16 = 2;
    pub const PS_DASHDOT: u16 = 3;
    pub const PS_DASHDOTDOT: u16 = 4;
    pub const PS_NULL: u16 = 5;
    pub const PS_INSIDEFRAME: u16 = 6;
    pub const PS_ALTERNATE: u16 = 7;
    pub const PS_ENDCAP_MASK: u16 = 0x0f00;
    pub const PS_ENDCAP_ROUND: u16 = 0x0000;
    pub const PS_ENDCAP_SQUARE: u16 = 0x0100;
    pub const PS_ENDCAP_FLAT: u16 = 0x0200;
    pub const PS_JOIN_MASK: u16 = 0xf000;
    pub const PS_JOIN_ROUND: u16 = 0x0000;
    pub const PS_JOIN_BEVEL: u16 = 0x1000;
    pub const PS_JOIN_MITER: u16 = 0x2000;
}

/// BrushStyle and HatchStyle values.
#[allow(dead_code)]
pub mod brush {
    pub const BS_SOLID: u16 = 0;
    pub const BS_NULL: u16 = 1;
    pub const BS_HOLLOW: u16 = BS_NULL;
    pub const BS_HATCHED: u16 = 2;
    pub const BS_PATTERN: u16 = 3;
    pub const BS_INDEXED: u16 = 4;
    pub const BS_DIBPATTERN: u16 = 5;
    pub const BS_DIBPATTERNPT: u16 = 6;
    pub const BS_PATTERN8X8: u16 = 7;
    pub const BS_DIBPATTERN8X8: u16 = 8;
    pub const HS_HORIZONTAL: u16 = 0;
    pub const HS_VERTICAL: u16 = 1;
    pub const HS_FDIAGONAL: u16 = 2;
    pub const HS_BDIAGONAL: u16 = 3;
    pub const HS_CROSS: u16 = 4;
    pub const HS_DIAGCROSS: u16 = 5;
}

#[allow(dead_code)]
pub mod fill_mode {
    pub const ALTERNATE: u16 = 1;
    pub const WINDING: u16 = 2;
}

#[allow(dead_code)]
pub mod text_align {
    pub const TA_NOUPDATECP: u16 = 0;
    pub const TA_UPDATECP: u16 = 1;
    pub const TA_LEFT: u16 = 0;
    pub const TA_RIGHT: u16 = 2;
    pub const TA_CENTER: u16 = 6;
    pub const TA_TOP: u16 = 0;
    pub const TA_BOTTOM: u16 = 8;
    pub const TA_BASELINE: u16 = 24;
    pub const TA_RTLREADING: u16 = 256;
    pub const HORIZONTAL_MASK: u16 = 6;
    pub const VERTICAL_MASK: u16 = 24;
}

#[allow(dead_code)]
pub mod ext_text_out {
    pub const ETO_OPAQUE: u16 = 0x0002;
    pub const ETO_CLIPPED: u16 = 0x0004;
    pub const ETO_GLYPH_INDEX: u16 = 0x0010;
    pub const ETO_RTLREADING: u16 = 0x0080;
    pub const ETO_PDY: u16 = 0x2000;
}

#[allow(dead_code)]
pub mod font_weight {
    pub const FW_DONTCARE: u16 = 0;
    pub const FW_THIN: u16 = 100;
    pub const FW_EXTRALIGHT: u16 = 200;
    pub const FW_LIGHT: u16 = 300;
    pub const FW_NORMAL: u16 = 400;
    pub const FW_MEDIUM: u16 = 500;
    pub const FW_SEMIBOLD: u16 = 600;
    pub const FW_BOLD: u16 = 700;
    pub const FW_EXTRABOLD: u16 = 800;
    pub const FW_HEAVY: u16 = 900;
}

#[allow(dead_code)]
pub mod bk_mode {
    pub const TRANSPARENT: u16 = 1;
    pub const OPAQUE: u16 = 2;
}

#[allow(dead_code)]
pub mod map_mode {
    pub const MM_TEXT: u16 = 1;
    pub const MM_LOMETRIC: u16 = 2;
    pub const MM_HIMETRIC: u16 = 3;
    pub const MM_LOENGLISH: u16 = 4;
    pub const MM_HIENGLISH: u16 = 5;
    pub const MM_TWIPS: u16 = 6;
    pub const MM_ISOTROPIC: u16 = 7;
    pub const MM_ANISOTROPIC: u16 = 8;
}

#[allow(dead_code)]
pub mod layout {
    pub const LTR: u16 = 0x0000;
    pub const RTL: u16 = 0x0001;
    pub const BITMAP_ORIENTATION_PRESERVED: u16 = 0x0008;
}

#[allow(dead_code)]
pub mod stock {
    pub const FLAG: u16 = 0x8000;
    pub const WHITE_BRUSH: u16 = 0;
    pub const LTGRAY_BRUSH: u16 = 1;
    pub const GRAY_BRUSH: u16 = 2;
    pub const DKGRAY_BRUSH: u16 = 3;
    pub const BLACK_BRUSH: u16 = 4;
    pub const NULL_BRUSH: u16 = 5;
    pub const WHITE_PEN: u16 = 6;
    pub const BLACK_PEN: u16 = 7;
    pub const NULL_PEN: u16 = 8;
    pub const OEM_FIXED_FONT: u16 = 10;
    pub const ANSI_FIXED_FONT: u16 = 11;
    pub const ANSI_VAR_FONT: u16 = 12;
    pub const SYSTEM_FONT: u16 = 13;
    pub const DEVICE_DEFAULT_FONT: u16 = 14;
    pub const DEFAULT_PALETTE: u16 = 15;
    pub const SYSTEM_FIXED_FONT: u16 = 16;
    pub const DEFAULT_GUI_FONT: u16 = 17;
}

/// Frequently occurring ternary raster operations.
#[allow(dead_code)]
pub mod rop3 {
    pub const BLACKNESS: u32 = 0x0000_0042;
    pub const DSTINVERT: u32 = 0x0055_0009;
    pub const PATINVERT: u32 = 0x005a_0049;
    pub const SRCCOPY: u32 = 0x00cc_0020;
    pub const PATCOPY: u32 = 0x00f0_0021;
    pub const WHITENESS: u32 = 0x00ff_0062;
}

#[cfg(test)]
mod tests {
    use super::record;

    #[test]
    fn corrected_record_values_match_ms_wmf() {
        assert_eq!(record::SET_PIXEL, 0x041f);
        assert_eq!(record::SET_TEXT_ALIGN, 0x012e);
        assert_eq!(record::SET_LAYOUT, 0x0149);
        assert_eq!(record::PAT_BLT, 0x061d);
        assert_ne!(record::SET_PIXEL, record::SET_VIEWPORT_ORG);
    }

    #[test]
    fn variable_function_high_byte_is_normalized() {
        assert_eq!(record::canonical(0x1224), record::POLYGON);
        assert_eq!(record::canonical(0x1737), record::SET_PALETTE_ENTRIES);
        assert_eq!(record::canonical(record::RECTANGLE), record::RECTANGLE);
    }

    #[test]
    fn record_catalog_contains_all_70_unique_record_types() {
        let mut records = record::ALL.to_vec();
        records.sort_unstable();
        records.dedup();
        assert_eq!(records.len(), 70);
    }
}
