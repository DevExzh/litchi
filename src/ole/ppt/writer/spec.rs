//! MS-PPT specification types and constants
//!
//! Reference: [MS-PPT] PowerPoint 97-2003 Binary File Format (.ppt)
//! https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt

// =============================================================================
// Slide Layout Types (MS-PPT 2.13.25 SlideLayoutType)
// =============================================================================

/// Slide layout geometry types from MS-PPT 2.13.25
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SlideLayoutType {
    /// SL_TitleSlide - Title slide layout
    TitleSlide = 0x0000,
    /// SL_TitleBody - Title and body layout (used by MainMaster)
    TitleBody = 0x0001,
    /// SL_MasterTitle - Master title layout
    MasterTitle = 0x0002,
    /// SL_TitleOnly - Title only layout
    TitleOnly = 0x0007,
    /// SL_TwoColumns - Two column layout
    TwoColumns = 0x0008,
    /// SL_TwoRows - Two row layout
    TwoRows = 0x0009,
    /// SL_ColumnTwoRows - Column with two rows
    ColumnTwoRows = 0x000A,
    /// SL_TwoRowsColumn - Two rows with column
    TwoRowsColumn = 0x000B,
    /// SL_TwoColumnsRow - Two columns with row
    TwoColumnsRow = 0x000C,
    /// SL_Blank - Blank slide layout
    Blank = 0x000D,
    /// SL_FourObjects - Four objects layout
    FourObjects = 0x000E,
    /// SL_BigObject - Big object layout
    BigObject = 0x000F,
    /// SL_VerticalTitleBody - Vertical title and body
    VerticalTitleBody = 0x0010,
    /// SL_VerticalTwoRows - Vertical two rows
    VerticalTwoRows = 0x0011,
}

// =============================================================================
// Placeholder Types (MS-PPT 2.13.23 PlaceholderType)
// =============================================================================

/// Placeholder types from MS-PPT 2.13.23
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PlaceholderType {
    /// PT_None - No placeholder
    None = 0x00,
    /// PT_MasterTitle - Master title placeholder
    MasterTitle = 0x01,
    /// PT_MasterBody - Master body placeholder
    MasterBody = 0x02,
    /// PT_MasterCenterTitle - Master center title
    MasterCenterTitle = 0x03,
    /// PT_MasterSubTitle - Master subtitle
    MasterSubTitle = 0x04,
    /// PT_MasterNotesSlideImage - Notes slide image
    MasterNotesSlideImage = 0x05,
    /// PT_MasterNotesBody - Notes body
    MasterNotesBody = 0x06,
    /// PT_MasterDate - Date placeholder
    MasterDate = 0x07,
    /// PT_MasterSlideNumber - Slide number placeholder
    MasterSlideNumber = 0x08,
    /// PT_MasterFooter - Footer placeholder
    MasterFooter = 0x09,
    /// PT_MasterHeader - Header placeholder
    MasterHeader = 0x0A,
}

// =============================================================================
// Slide Flags (MS-PPT 2.4.7 SlideAtom)
// =============================================================================

/// Slide flags from SlideAtom (MS-PPT 2.4.7)
pub mod slide_flags {
    /// fMasterObjects - Follow master objects
    pub const MASTER_OBJECTS: u16 = 0x0001;
    /// fMasterScheme - Follow master color scheme
    pub const MASTER_SCHEME: u16 = 0x0002;
    /// fMasterBackground - Follow master background
    pub const MASTER_BACKGROUND: u16 = 0x0004;
    /// Default: all master flags enabled
    pub const DEFAULT: u16 = MASTER_OBJECTS | MASTER_SCHEME | MASTER_BACKGROUND;
}

// =============================================================================
// Color Scheme (MS-PPT 2.4.17 ColorSchemeAtom)
// =============================================================================

/// Color scheme with 8 RGBX colors (MS-PPT 2.4.17)
///
/// The color scheme defines the palette used by slides.
/// Each color is stored as RGBX (4 bytes: R, G, B, unused).
#[derive(Debug, Clone, Copy)]
pub struct ColorScheme {
    /// Background color
    pub background: u32,
    /// Text and lines color
    pub text_and_lines: u32,
    /// Shadow color
    pub shadow: u32,
    /// Title text color
    pub title_text: u32,
    /// Fill color
    pub fill: u32,
    /// Accent color
    pub accent: u32,
    /// Accent and hyperlink color
    pub accent_and_hyperlink: u32,
    /// Accent and followed hyperlink color
    pub accent_and_followed_hyperlink: u32,
}

impl ColorScheme {
    /// POI's default color scheme for slides
    pub const POI_DEFAULT: Self = Self {
        background: 0x00FFFFFF,
        text_and_lines: 0x00000000,
        shadow: 0x00808080,
        title_text: 0x00000000,
        fill: 0x00996630,
        accent: 0x00CC9963,
        accent_and_hyperlink: 0x00FFCC66,
        accent_and_followed_hyperlink: 0x00B256AE,
    };

    /// Convert to 32-byte array for writing
    pub fn to_bytes(&self) -> [u8; 32] {
        let mut data = [0u8; 32];
        data[0..4].copy_from_slice(&self.background.to_le_bytes());
        data[4..8].copy_from_slice(&self.text_and_lines.to_le_bytes());
        data[8..12].copy_from_slice(&self.shadow.to_le_bytes());
        data[12..16].copy_from_slice(&self.title_text.to_le_bytes());
        data[16..20].copy_from_slice(&self.fill.to_le_bytes());
        data[20..24].copy_from_slice(&self.accent.to_le_bytes());
        data[24..28].copy_from_slice(&self.accent_and_hyperlink.to_le_bytes());
        data[28..32].copy_from_slice(&self.accent_and_followed_hyperlink.to_le_bytes());
        data
    }
}

// =============================================================================
// PPT10 Binary Tag (PowerPoint 2002+ features)
// =============================================================================

/// PPT10 tag string - marks PowerPoint 2002 (XP) and later features
pub struct Ppt10Tag;

impl Ppt10Tag {
    /// The tag identifier string
    pub const TAG_STRING: &'static str = "___PPT10";

    /// Convert to UTF-16LE bytes for writing
    pub fn to_bytes() -> [u8; 16] {
        let mut data = [0u8; 16];
        for (i, ch) in Self::TAG_STRING.encode_utf16().enumerate() {
            let bytes = ch.to_le_bytes();
            data[i * 2] = bytes[0];
            data[i * 2 + 1] = bytes[1];
        }
        data
    }
}

/// PP10SlideBinaryTagExtension - Binary tag data structure (MS-PPT 2.4.22.4)
///
/// Contains extended data for PowerPoint 2002+ features.
#[derive(Debug, Clone, Copy)]
pub struct BinaryTagData {
    /// Reserved/padding field
    pub reserved: u16,
    /// Tag type identifier (0x2EEB for MainMaster, 0x040D for Slide/DocInfo)
    pub tag_type: u16,
    /// Length of the following data (typically 8)
    pub data_length: u32,
    /// First data word (timestamp low bits or flags)
    pub data_word1: u32,
    /// Second data word (timestamp high bits or flags)
    pub data_word2: u32,
}

impl BinaryTagData {
    /// Tag type for MainMaster ProgTags
    pub const TAG_TYPE_MAIN_MASTER: u16 = 0x2EEB;
    /// Tag type for Slide/DocInfo ProgTags
    pub const TAG_TYPE_SLIDE: u16 = 0x040D;

    /// MainMaster binary tag data (contains GUID-like timestamp)
    pub const MAIN_MASTER: Self = Self {
        reserved: 0,
        tag_type: Self::TAG_TYPE_MAIN_MASTER,
        data_length: 8,
        data_word1: 0x01C2_F3F7, // GUID-like data from POI
        data_word2: 0x4F2D_3670,
    };

    /// Slide binary tag data
    pub const SLIDE: Self = Self {
        reserved: 0,
        tag_type: Self::TAG_TYPE_SLIDE,
        data_length: 8,
        data_word1: 0,
        data_word2: 0,
    };

    /// DocInfo binary tag data (contains zoom factor)
    pub const DOCINFO: Self = Self {
        reserved: 0,
        tag_type: Self::TAG_TYPE_SLIDE,
        data_length: 8,
        data_word1: 0x0000_C000, // Zoom factor (192 = 0xC0 as fixed point)
        data_word2: 0x0000_C000,
    };

    pub fn to_bytes(&self) -> [u8; 16] {
        let mut data = [0u8; 16];
        data[0..2].copy_from_slice(&self.reserved.to_le_bytes());
        data[2..4].copy_from_slice(&self.tag_type.to_le_bytes());
        data[4..8].copy_from_slice(&self.data_length.to_le_bytes());
        data[8..12].copy_from_slice(&self.data_word1.to_le_bytes());
        data[12..16].copy_from_slice(&self.data_word2.to_le_bytes());
        data
    }
}

// =============================================================================
// UserEditAtom constants (MS-PPT 2.4.16)
// =============================================================================

/// PPT version field for UserEditAtom
/// This opaque value is from POI's empty.ppt
pub const PPT_VERSION: u32 = 0x0300106D;

/// Default lastViewedSlideID (256 = first slide)
pub const DEFAULT_LAST_VIEWED_SLIDE_ID: u32 = 256;

/// Default lastViewType (1 = slide view)
pub const DEFAULT_LAST_VIEW_TYPE: u16 = 1;

/// Padword value from POI empty.ppt
pub const USER_EDIT_PADWORD: u16 = 0x07B9;

// =============================================================================
// Escher Property IDs (MS-ODRAW 2.3.7)
// =============================================================================

/// Escher property IDs from MS-ODRAW specification
pub mod escher_prop {
    /// Fill color property
    pub const FILL_COLOR: u16 = 0x0181;
    /// Fill back color property
    pub const FILL_BACK_COLOR: u16 = 0x0183;
    /// Fill blip property (with fComplex flag)
    pub const FILL_BLIP: u16 = 0x4186;
    /// Line style boolean properties
    pub const LINE_STYLE_BOOL: u16 = 0x01BF;
    /// Line color property
    pub const LINE_COLOR: u16 = 0x01C0;
    /// Line blip property (with fComplex flag)
    pub const LINE_BLIP: u16 = 0x41C5;
    /// Shape boolean properties
    pub const SHAPE_BOOL: u16 = 0x01FF;
    /// Shadow color property
    pub const SHADOW_COLOR: u16 = 0x0201;
}

/// Escher scheme color values (MS-ODRAW 2.2.2)
pub mod escher_color {
    /// Use scheme color - ORed with color index
    pub const USE_SCHEME: u32 = 0x08000000;
    /// Scheme color index: fill
    pub const SCHEME_FILL: u32 = 0x04;
    /// Scheme color index: line
    pub const SCHEME_LINE: u32 = 0x01;
    /// Scheme color index: shadow
    pub const SCHEME_SHADOW: u32 = 0x02;
    /// Scheme color index: background
    pub const SCHEME_BACKGROUND: u32 = 0x00;
}

// =============================================================================
// Escher Record Types (MS-ODRAW 2.1.4)
// =============================================================================

/// Escher record types from MS-ODRAW specification
pub mod escher_type {
    /// DggContainer - Drawing group container
    pub const DGG_CONTAINER: u16 = 0xF000;
    /// Dgg - Drawing group data
    pub const DGG: u16 = 0xF006;
    /// DgContainer - Drawing container
    pub const DG_CONTAINER: u16 = 0xF002;
    /// Dg - Drawing data
    pub const DG: u16 = 0xF008;
    /// SpgrContainer - Shape group container
    pub const SPGR_CONTAINER: u16 = 0xF003;
    /// SpContainer - Shape container
    pub const SP_CONTAINER: u16 = 0xF004;
    /// Spgr - Shape group data
    pub const SPGR: u16 = 0xF009;
    /// Sp - Shape data
    pub const SP: u16 = 0xF00A;
    /// Opt - Property table
    pub const OPT: u16 = 0xF00B;
    /// ClientAnchor - Anchor data
    pub const CLIENT_ANCHOR: u16 = 0xF010;
    /// ClientData - Client data
    pub const CLIENT_DATA: u16 = 0xF011;
    /// SplitMenuColors - Menu colors
    pub const SPLIT_MENU_COLORS: u16 = 0xF11E;
}

/// Escher shape types (MS-ODRAW 2.2.15 MSOSPT)
pub mod escher_shape {
    /// msosptRectangle
    pub const RECTANGLE: u16 = 1;
    /// msosptRoundRectangle
    pub const ROUND_RECTANGLE: u16 = 2;
    /// msosptEllipse
    pub const ELLIPSE: u16 = 3;
    /// msosptTextBox
    pub const TEXT_BOX: u16 = 202;
}

/// Escher shape flags (MS-ODRAW 2.2.16)
pub mod escher_flags {
    /// fGroup - This is a group shape
    pub const GROUP: u32 = 0x0001;
    /// fChild - This is a child shape
    pub const CHILD: u32 = 0x0002;
    /// fPatriarch - This is the patriarch (root) shape
    pub const PATRIARCH: u32 = 0x0004;
    /// fDeleted - Shape is deleted
    pub const DELETED: u32 = 0x0008;
    /// fOleShape - Shape is an OLE object
    pub const OLE_SHAPE: u32 = 0x0010;
    /// fHaveMaster - Shape has a master
    pub const HAVE_MASTER: u32 = 0x0020;
    /// fFlipH - Flip horizontally
    pub const FLIP_H: u32 = 0x0040;
    /// fFlipV - Flip vertically
    pub const FLIP_V: u32 = 0x0080;
    /// fConnector - Shape is a connector
    pub const CONNECTOR: u32 = 0x0100;
    /// fHaveAnchor - Shape has an anchor
    pub const HAVE_ANCHOR: u32 = 0x0200;
    /// fBackground - Shape is a background
    pub const BACKGROUND: u32 = 0x0400;
    /// fHaveSpt - Shape has a shape type
    pub const HAVE_SPT: u32 = 0x0800;
}

// =============================================================================
// MainMaster SlideAtom Placeholders (MS-PPT 2.4.7)
// =============================================================================

/// MainMaster placeholder types array for SlideAtom
/// Order: MasterTitle, MasterBody, MasterDate, MasterFooter, MasterSlideNumber, None, None, None
pub const MAIN_MASTER_PLACEHOLDERS: [u8; 8] = [
    PlaceholderType::MasterTitle as u8,
    PlaceholderType::MasterBody as u8,
    PlaceholderType::MasterDate as u8,
    PlaceholderType::MasterFooter as u8,
    PlaceholderType::MasterSlideNumber as u8,
    PlaceholderType::None as u8,
    PlaceholderType::None as u8,
    PlaceholderType::None as u8,
];

/// MainMaster SlideAtom reserved field value from POI
pub const MAIN_MASTER_SLIDE_ATOM_RESERVED: u16 = 0x0013;

// =============================================================================
// MainMaster Color Schemes (12 pre-defined schemes from POI empty.ppt)
// =============================================================================

/// Helper to create RGBX color from RGB values
pub const fn rgb(r: u8, g: u8, b: u8) -> u32 {
    (r as u32) | ((g as u32) << 8) | ((b as u32) << 16)
}

/// Pre-defined color schemes for MainMaster (MS-PPT 2.4.17)
pub mod color_schemes {
    use super::{ColorScheme, rgb};

    /// Scheme 0: Default light (cyan/purple accent)
    pub const DEFAULT_LIGHT: ColorScheme = ColorScheme {
        background: rgb(0xFF, 0xFF, 0xFF),                    // white
        text_and_lines: rgb(0x00, 0x00, 0x00),                // black
        shadow: rgb(0x80, 0x80, 0x80),                        // gray
        title_text: rgb(0x00, 0x00, 0x00),                    // black
        fill: rgb(0xBB, 0xE0, 0xE3),                          // light cyan
        accent: rgb(0x33, 0x33, 0x99),                        // purple
        accent_and_hyperlink: rgb(0x00, 0x99, 0x99),          // teal
        accent_and_followed_hyperlink: rgb(0x99, 0xCC, 0x00), // yellow-green
    };

    /// Scheme 1: Golden accent
    pub const GOLDEN: ColorScheme = ColorScheme {
        background: rgb(0xFF, 0xFF, 0xFF),
        text_and_lines: rgb(0x00, 0x00, 0x00),
        shadow: rgb(0x96, 0x96, 0x96),
        title_text: rgb(0x00, 0x00, 0x00),
        fill: rgb(0xFB, 0xDF, 0x53),                          // gold
        accent: rgb(0xFF, 0x99, 0x66),                        // peach
        accent_and_hyperlink: rgb(0xCC, 0x33, 0x00),          // red-orange
        accent_and_followed_hyperlink: rgb(0x99, 0x66, 0x00), // brown
    };

    /// Scheme 2: Blue accent
    pub const BLUE_ACCENT: ColorScheme = ColorScheme {
        background: rgb(0xFF, 0xFF, 0xFF),
        text_and_lines: rgb(0x00, 0x00, 0x00),
        shadow: rgb(0x80, 0x80, 0x80),
        title_text: rgb(0x00, 0x00, 0x00),
        fill: rgb(0x99, 0xCC, 0xFF),                 // light blue
        accent: rgb(0xCC, 0xCC, 0xFF),               // lavender
        accent_and_hyperlink: rgb(0x33, 0x33, 0xCC), // blue
        accent_and_followed_hyperlink: rgb(0xAF, 0x67, 0xFF), // purple
    };

    /// Scheme 3: Mint green
    pub const MINT: ColorScheme = ColorScheme {
        background: rgb(0xDE, 0xF6, 0xF1), // mint background
        text_and_lines: rgb(0x00, 0x00, 0x00),
        shadow: rgb(0x96, 0x96, 0x96),
        title_text: rgb(0x00, 0x00, 0x00),
        fill: rgb(0xFF, 0xFF, 0xFF),
        accent: rgb(0x8D, 0xC6, 0xFF),
        accent_and_hyperlink: rgb(0x00, 0x66, 0xCC),
        accent_and_followed_hyperlink: rgb(0x00, 0xA8, 0x00),
    };

    /// Scheme 4: Cream/coral
    pub const CREAM: ColorScheme = ColorScheme {
        background: rgb(0xFF, 0xFF, 0xD9), // cream background
        text_and_lines: rgb(0x00, 0x00, 0x00),
        shadow: rgb(0x77, 0x77, 0x77),
        title_text: rgb(0x00, 0x00, 0x00),
        fill: rgb(0xFF, 0xFF, 0xF7),
        accent: rgb(0x33, 0xCC, 0xCC),
        accent_and_hyperlink: rgb(0xFF, 0x50, 0x50), // coral
        accent_and_followed_hyperlink: rgb(0xFF, 0x99, 0x00),
    };

    /// Scheme 5: Teal dark
    pub const TEAL_DARK: ColorScheme = ColorScheme {
        background: rgb(0x00, 0x80, 0x80),     // teal background
        text_and_lines: rgb(0xFF, 0xFF, 0xFF), // white text
        shadow: rgb(0x00, 0x5A, 0x58),
        title_text: rgb(0xFF, 0xFF, 0x99), // yellow
        fill: rgb(0x00, 0x64, 0x62),
        accent: rgb(0x6D, 0x6F, 0xC7),
        accent_and_hyperlink: rgb(0x00, 0xFF, 0xFF),
        accent_and_followed_hyperlink: rgb(0x00, 0xFF, 0x00),
    };

    /// Scheme 6: Maroon dark
    pub const MAROON_DARK: ColorScheme = ColorScheme {
        background: rgb(0x80, 0x00, 0x00), // maroon background
        text_and_lines: rgb(0xFF, 0xFF, 0xFF),
        shadow: rgb(0x5C, 0x1F, 0x00),
        title_text: rgb(0xDF, 0xD2, 0x93),
        fill: rgb(0xCC, 0x33, 0x00),
        accent: rgb(0xBE, 0x79, 0x60),
        accent_and_hyperlink: rgb(0xFF, 0xFF, 0x99),
        accent_and_followed_hyperlink: rgb(0xD3, 0xA2, 0x19),
    };

    /// Scheme 7: Navy dark
    pub const NAVY_DARK: ColorScheme = ColorScheme {
        background: rgb(0x00, 0x00, 0x99), // navy background
        text_and_lines: rgb(0xFF, 0xFF, 0xFF),
        shadow: rgb(0x00, 0x33, 0x66),
        title_text: rgb(0xCC, 0xFF, 0xFF),
        fill: rgb(0x33, 0x66, 0xCC),
        accent: rgb(0x00, 0xB0, 0x00),
        accent_and_hyperlink: rgb(0x66, 0xCC, 0xFF),
        accent_and_followed_hyperlink: rgb(0xFF, 0xE7, 0x01),
    };

    /// Scheme 8: Black dark
    pub const BLACK_DARK: ColorScheme = ColorScheme {
        background: rgb(0x00, 0x00, 0x00), // black background
        text_and_lines: rgb(0xFF, 0xFF, 0xFF),
        shadow: rgb(0x33, 0x66, 0x99),
        title_text: rgb(0xE3, 0xEB, 0xF1),
        fill: rgb(0x00, 0x33, 0x99),
        accent: rgb(0x46, 0x8A, 0x4B),
        accent_and_hyperlink: rgb(0x66, 0xCC, 0xFF),
        accent_and_followed_hyperlink: rgb(0xF0, 0xE5, 0x00),
    };

    /// Scheme 9: Olive
    pub const OLIVE: ColorScheme = ColorScheme {
        background: rgb(0x68, 0x6B, 0x5D), // olive background
        text_and_lines: rgb(0xFF, 0xFF, 0xFF),
        shadow: rgb(0x77, 0x77, 0x77),
        title_text: rgb(0xD1, 0xD1, 0xCB),
        fill: rgb(0x90, 0x90, 0x82),
        accent: rgb(0x80, 0x9E, 0xA8),
        accent_and_hyperlink: rgb(0xFF, 0xCC, 0x66),
        accent_and_followed_hyperlink: rgb(0xE9, 0xDC, 0xB9),
    };

    /// Scheme 10: Purple
    pub const PURPLE: ColorScheme = ColorScheme {
        background: rgb(0x66, 0x66, 0x99), // purple background
        text_and_lines: rgb(0xFF, 0xFF, 0xFF),
        shadow: rgb(0x3E, 0x3E, 0x5C),
        title_text: rgb(0xFF, 0xFF, 0xFF),
        fill: rgb(0x60, 0x59, 0x7B),
        accent: rgb(0x66, 0x66, 0xFF),
        accent_and_hyperlink: rgb(0x99, 0xCC, 0xFF),
        accent_and_followed_hyperlink: rgb(0xFF, 0xFF, 0x99),
    };

    /// Scheme 11: Brown
    pub const BROWN: ColorScheme = ColorScheme {
        background: rgb(0x52, 0x3E, 0x26), // brown background
        text_and_lines: rgb(0xFF, 0xFF, 0xFF),
        shadow: rgb(0x2D, 0x20, 0x15),
        title_text: rgb(0xDF, 0xC0, 0x8D),
        fill: rgb(0x8C, 0x7B, 0x70),
        accent: rgb(0x8F, 0x5F, 0x2F),
        accent_and_hyperlink: rgb(0xCC, 0xB4, 0x00),
        accent_and_followed_hyperlink: rgb(0x8C, 0x9E, 0xA0),
    };

    /// All 12 MainMaster color schemes in order
    pub const ALL: [ColorScheme; 12] = [
        DEFAULT_LIGHT,
        GOLDEN,
        BLUE_ACCENT,
        MINT,
        CREAM,
        TEAL_DARK,
        MAROON_DARK,
        NAVY_DARK,
        BLACK_DARK,
        OLIVE,
        PURPLE,
        BROWN,
    ];
}
