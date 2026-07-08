//! WMF constants and enumerations
//!
//! Defines all WMF record types, pen styles, brush styles, and other constants
//! used in Windows Metafile format.

/// WMF record function codes
#[allow(dead_code)]
pub mod record {
    // State records
    pub const SAVE_DC: u16 = 0x001E;
    pub const RESTORE_DC: u16 = 0x0127;
    pub const SET_BK_COLOR: u16 = 0x0201;
    pub const SET_BK_MODE: u16 = 0x0102;
    pub const SET_MAP_MODE: u16 = 0x0103;
    pub const SET_ROP2: u16 = 0x0104;
    pub const SET_REL_ABS: u16 = 0x0105;
    pub const SET_POLY_FILL_MODE: u16 = 0x0106;
    pub const SET_STRETCH_BLT_MODE: u16 = 0x0107;
    pub const SET_TEXT_CHAR_EXTRA: u16 = 0x0108;
    pub const SET_TEXT_COLOR: u16 = 0x0209;
    pub const SET_TEXT_JUSTIFICATION: u16 = 0x020A;
    pub const SET_WINDOW_ORG: u16 = 0x020B;
    pub const SET_WINDOW_EXT: u16 = 0x020C;
    pub const SET_VIEWPORT_ORG: u16 = 0x020D;
    pub const SET_VIEWPORT_EXT: u16 = 0x020E;
    pub const OFFSET_WINDOW_ORG: u16 = 0x020F;
    pub const SCALE_WINDOW_EXT: u16 = 0x0410;
    pub const OFFSET_VIEWPORT_ORG: u16 = 0x0211;
    pub const SCALE_VIEWPORT_EXT: u16 = 0x0412;
    pub const SET_PIXEL_V: u16 = 0x020D;

    // Drawing records
    pub const LINE_TO: u16 = 0x0213;
    pub const MOVE_TO: u16 = 0x0214;
    pub const POLYGON: u16 = 0x0324;
    pub const POLYLINE: u16 = 0x0325;
    pub const RECTANGLE: u16 = 0x041B;
    pub const ROUND_RECT: u16 = 0x061C;
    pub const ELLIPSE: u16 = 0x0418;
    pub const ARC: u16 = 0x0817;
    pub const PIE: u16 = 0x081A;
    pub const CHORD: u16 = 0x0830;
    pub const POLYPOLYGON: u16 = 0x0538;
    pub const POLYGON16: u16 = 0x0324;
    pub const POLYLINE16: u16 = 0x0325;

    // Text records
    pub const TEXT_OUT: u16 = 0x0521;
    pub const EXT_TEXT_OUT: u16 = 0x0A32;
    pub const DRAW_TEXT: u16 = 0x062F;

    // Object records
    pub const CREATE_PEN_INDIRECT: u16 = 0x02FA;
    pub const CREATE_BRUSH_INDIRECT: u16 = 0x02FC;
    pub const CREATE_FONT_INDIRECT: u16 = 0x02FB;
    pub const CREATE_PALETTE: u16 = 0x00F7;
    pub const CREATE_REGION: u16 = 0x06FF;
    pub const CREATE_PATTERN_BRUSH: u16 = 0x01F9;
    pub const CREATE_DIB_PATTERN_BRUSH: u16 = 0x0142;
    pub const SELECT_OBJECT: u16 = 0x012D;
    pub const DELETE_OBJECT: u16 = 0x01F0;
    pub const SELECT_PALETTE: u16 = 0x0234;
    pub const REALIZE_PALETTE: u16 = 0x0035;
    pub const ANIMATE_PALETTE: u16 = 0x0436;
    pub const SET_PALETTE_ENTRIES: u16 = 0x0037;
    pub const RESIZE_PALETTE: u16 = 0x0139;

    // Bitmap records
    pub const BIT_BLT: u16 = 0x0922;
    pub const STRETCH_BLT: u16 = 0x0B23;
    pub const DIB_BIT_BLT: u16 = 0x0940;
    pub const DIB_STRETCH_BLT: u16 = 0x0B41;
    pub const SET_DIB_TO_DEV: u16 = 0x0D33;
    pub const STRETCH_DIB: u16 = 0x0F43;
    pub const DIB_CREATE_PATTERN_BRUSH: u16 = 0x0142;

    // Clipping records
    pub const EXCLUDE_CLIP_RECT: u16 = 0x0415;
    pub const INTERSECT_CLIP_RECT: u16 = 0x0416;
    pub const SELECT_CLIP_REGION: u16 = 0x012C;
    pub const OFFSET_CLIP_RGN: u16 = 0x0220;

    // Fill records
    pub const FLOOD_FILL: u16 = 0x0419;
    pub const EXT_FLOOD_FILL: u16 = 0x0548;
    pub const FILL_REGION: u16 = 0x0228;
    pub const FRAME_REGION: u16 = 0x0429;
    pub const INVERT_REGION: u16 = 0x012A;
    pub const PAINT_REGION: u16 = 0x012B;

    // Control records
    pub const EOF: u16 = 0x0000;
    pub const SET_MAPPER_FLAGS: u16 = 0x0231;
    pub const ESCAPE: u16 = 0x0626;
}

/// Pen style constants
#[allow(dead_code)]
pub mod pen {
    // Base styles (lower 4 bits)
    pub const PS_SOLID: u16 = 0;
    pub const PS_DASH: u16 = 1;
    pub const PS_DOT: u16 = 2;
    pub const PS_DASHDOT: u16 = 3;
    pub const PS_DASHDOTDOT: u16 = 4;
    pub const PS_NULL: u16 = 5;
    pub const PS_INSIDEFRAME: u16 = 6;
    pub const PS_ALTERNATE: u16 = 7;

    // End cap styles (bits 8-11)
    pub const PS_ENDCAP_ROUND: u16 = 0x0000;
    pub const PS_ENDCAP_SQUARE: u16 = 0x0100;
    pub const PS_ENDCAP_FLAT: u16 = 0x0200;

    // Join styles (bits 12-15)
    pub const PS_JOIN_ROUND: u16 = 0x0000;
    pub const PS_JOIN_BEVEL: u16 = 0x1000;
    pub const PS_JOIN_MITER: u16 = 0x2000;
}

/// Brush style constants
#[allow(dead_code)]
pub mod brush {
    pub const BS_SOLID: u16 = 0;
    pub const BS_NULL: u16 = 1;
    pub const BS_HOLLOW: u16 = 1; // Same as BS_NULL
    pub const BS_HATCHED: u16 = 2;
    pub const BS_PATTERN: u16 = 3;
    pub const BS_INDEXED: u16 = 4;
    pub const BS_DIBPATTERN: u16 = 5;
    pub const BS_DIBPATTERNPT: u16 = 6;
    pub const BS_PATTERN8X8: u16 = 7;
    pub const BS_DIBPATTERN8X8: u16 = 8;

    // Hatch styles (for BS_HATCHED)
    pub const HS_HORIZONTAL: u16 = 0;
    pub const HS_VERTICAL: u16 = 1;
    pub const HS_FDIAGONAL: u16 = 2;
    pub const HS_BDIAGONAL: u16 = 3;
    pub const HS_CROSS: u16 = 4;
    pub const HS_DIAGCROSS: u16 = 5;
}

/// Polygon fill modes
#[allow(dead_code)]
pub mod fill_mode {
    pub const ALTERNATE: u16 = 1; // Even-odd fill (SVG evenodd)
    pub const WINDING: u16 = 2; // Non-zero winding (SVG nonzero)
}

/// Text alignment modes
#[allow(dead_code)]
pub mod text_align {
    pub const TA_LEFT: u16 = 0;
    pub const TA_CENTER: u16 = 6;
    pub const TA_RIGHT: u16 = 2;
    pub const TA_TOP: u16 = 0;
    pub const TA_BOTTOM: u16 = 8;
    pub const TA_BASELINE: u16 = 24;
    pub const TA_UPDATECP: u16 = 1;
}

/// Font weights
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

/// Background modes
#[allow(dead_code)]
pub mod bk_mode {
    pub const TRANSPARENT: u16 = 1;
    pub const OPAQUE: u16 = 2;
}

/// Mapping modes
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
