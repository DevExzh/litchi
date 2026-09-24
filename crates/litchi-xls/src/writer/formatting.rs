//! XLS cell formatting (XF records, fonts, fills, borders)
//!
//! This module implements BIFF8 formatting records for Excel 97-2003 files.
//! Based on Microsoft's "[MS-XLS]" specification and Apache POI's implementation.
//!
//! # Key Structures
//!
//! - **XF (Extended Format)**: Cell format combining font, fill, border, and alignment
//! - **FONT**: Font definition (name, size, color, style)
//! - **FORMAT**: Number format definition
//! - **PALETTE**: Color palette

use super::super::{Error, Result};
use super::string_limits::{
    FONT_NAME_UNITS, NUMBER_FORMAT_UNITS, ensure_utf16_len_within, record_len, u8_len, utf16_len,
};
use std::collections::HashMap;
use std::io::Write;

/// `Font` records a workbook can address: `FontIndex` skips 4 and MUST be
/// at most 1022 ([MS-XLS] 2.5.129), which names 1022 records. The reader
/// refuses a font it cannot address.
pub(crate) const MAX_FONTS: usize = 1022;

/// `Format` records the globals grammar allows, `8*218Format` ([MS-XLS]
/// 2.1.7.20.1), less the eight locale-dependent built-ins this writer always
/// emits. The reader refuses a 219th `Format` record.
pub(crate) const MAX_CUSTOM_NUMBER_FORMATS: usize = 218 - 8;

/// XF records a 16-bit `XFIndex` ([MS-XLS] 2.5.282) can address; the reader
/// refuses more.
pub(crate) const MAX_XF_RECORDS: usize = 65_536;

/// XF records when an `XFCRC` record is written, as it is for `XFExt`
/// records and for pivot tables: `XFCRC.cxfs` is 16 through 4050 ([MS-XLS]
/// 2.4.354), and the reader refuses anything else.
pub(crate) const MAX_XF_RECORDS_WITH_XFCRC: usize = 4050;

/// The fifteen built-in style XFs, the default cell XF and the five built-in
/// number-format style XFs that precede user cell XFs.
const FIXED_XF_RECORDS: usize = 15 + 1 + 5;

/// The first XF index the pivot-table XFs may take.
const PIVOT_XF_START_INDEX: usize = 64;

/// The XF records pivot-table formatting appends.
const PIVOT_XF_RECORDS: usize = 3;

/// Cell formats a writer can hold (the default one included): their XFs, the
/// fixed XFs before them and the pivot XFs after them fill the XF index
/// space. The pivot XFs are reserved even without a pivot table, so adding a
/// pivot table can never overflow it.
pub(crate) const MAX_CELL_FORMATS: usize = MAX_XF_RECORDS - PIVOT_XF_RECORDS - FIXED_XF_RECORDS + 1;

/// The XF records [`FormattingManager::write_formats`] emits for `formats`
/// cell formats (the default one included), with or without pivot XFs.
pub(crate) const fn xf_record_count_for(formats: usize, pivot: bool) -> usize {
    let base = FIXED_XF_RECORDS + formats.saturating_sub(1);
    if !pivot {
        base
    } else if base > PIVOT_XF_START_INDEX {
        base + PIVOT_XF_RECORDS
    } else {
        PIVOT_XF_START_INDEX + PIVOT_XF_RECORDS
    }
}

#[derive(Debug, Clone, Copy)]
pub struct PivotXfIndices {
    pub header_accent: u16,
    pub row_label: u16,
    pub value: u16,
}

/// Font weight constants
pub const FONT_WEIGHT_NORMAL: u16 = 400;
pub const FONT_WEIGHT_BOLD: u16 = 700;

/// Default color indices
pub const COLOR_BLACK: u16 = 0x08;
pub const COLOR_WHITE: u16 = 0x09;
pub const COLOR_RED: u16 = 0x0A;
pub const COLOR_AUTOMATIC: u16 = 0x7FFF;

/// Horizontal alignment
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HorizontalAlignment {
    General = 0,
    Left = 1,
    Center = 2,
    Right = 3,
    Fill = 4,
    Justify = 5,
    CenterAcrossSelection = 6,
}

/// Built-in number format strings as defined by BIFF8 / Excel.
///
/// These are taken from Apache POI's `BuiltinFormats` table so that
/// format indices used in `ExtendedFormat.format_index` match POI and
/// Excel expectations.
const BUILTIN_NUMBER_FORMATS: [&str; 50] = [
    "General",                              // 0x00
    "0",                                    // 0x01
    "0.00",                                 // 0x02
    "#,##0",                                // 0x03
    "#,##0.00",                             // 0x04
    "\"$\"#,##0_);(\"$\"#,##0)",            // 0x05
    "\"$\"#,##0_);[Red](\"$\"#,##0)",       // 0x06
    "\"$\"#,##0.00_);(\"$\"#,##0.00)",      // 0x07
    "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)", // 0x08
    "0%",                                   // 0x09
    "0.00%",                                // 0x0A
    "0.00E+00",                             // 0x0B
    "# ?/?",                                // 0x0C
    "# ??/??",                              // 0x0D
    "m/d/yy",                               // 0x0E
    "d-mmm-yy",                             // 0x0F
    "d-mmm",                                // 0x10
    "mmm-yy",                               // 0x11
    "h:mm AM/PM",                           // 0x12
    "h:mm:ss AM/PM",                        // 0x13
    "h:mm",                                 // 0x14
    "h:mm:ss",                              // 0x15
    "m/d/yy h:mm",                          // 0x16
    // 0x17 - 0x24 reserved for international and undocumented
    "reserved-0x17",              // 0x17
    "reserved-0x18",              // 0x18
    "reserved-0x19",              // 0x19
    "reserved-0x1A",              // 0x1A
    "reserved-0x1B",              // 0x1B
    "reserved-0x1C",              // 0x1C
    "reserved-0x1D",              // 0x1D
    "reserved-0x1E",              // 0x1E
    "reserved-0x1F",              // 0x1F
    "reserved-0x20",              // 0x20
    "reserved-0x21",              // 0x21
    "reserved-0x22",              // 0x22
    "reserved-0x23",              // 0x23
    "reserved-0x24",              // 0x24
    "#,##0_);(#,##0)",            // 0x25
    "#,##0_);[Red](#,##0)",       // 0x26
    "#,##0.00_);(#,##0.00)",      // 0x27
    "#,##0.00_);[Red](#,##0.00)", // 0x28
    "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
    "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)",
    "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
    "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)",
    "mm:ss",     // 0x2D
    "[h]:mm:ss", // 0x2E
    "mm:ss.0",   // 0x2F
    "##0.0E+0",  // 0x30
    "@",         // 0x31 (text)
];

/// First user-defined number format index in BIFF8 / Excel.
const FIRST_USER_DEFINED_NUMBER_FORMAT_INDEX: u16 = 164;

/// Look up the BIFF built-in number format index for a given pattern.
///
/// This mirrors Apache POI's `BuiltinFormats.getBuiltinFormat(String)`
/// in a simplified form and is used by the formatting manager to avoid
/// creating duplicate custom FORMAT records for built-in patterns.
/// Refuses a number format a `Format` record cannot hold: 1 through 255 UTF-16
/// code units ([MS-XLS] 2.4.126; the reader refuses anything else).
pub(crate) fn validate_number_format(pattern: &str) -> Result<()> {
    if pattern.is_empty() {
        return Err(Error::InvalidData(
            "number format must not be empty".to_string(),
        ));
    }
    ensure_utf16_len_within(pattern, NUMBER_FORMAT_UNITS, "number format")
}

fn builtin_number_format_index(pattern: &str) -> Option<u16> {
    BUILTIN_NUMBER_FORMATS
        .iter()
        .position(|&p| p == pattern)
        .map(crate::utils::truncate_usize_to_u16)
}

/// Vertical alignment
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlignment {
    Top = 0,
    Center = 1,
    Bottom = 2,
    Justify = 3,
}

/// Border style
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum BorderStyle {
    #[default]
    None = 0,
    Thin = 1,
    Medium = 2,
    Dashed = 3,
    Dotted = 4,
    Thick = 5,
    Double = 6,
    Hair = 7,
}

/// Fill pattern
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FillPattern {
    None = 0,
    Solid = 1,
    MediumGray = 2,
    DarkGray = 3,
    LightGray = 4,
    DarkHorizontal = 5,
    DarkVertical = 6,
    DarkDown = 7,
    DarkUp = 8,
    DarkGrid = 9,
    DarkTrellis = 10,
}

/// Font definition
#[derive(Debug, Clone)]
pub struct Font {
    /// Font height in twips (1/20 of a point)
    pub height: u16,
    /// Font weight (400 = normal, 700 = bold)
    pub weight: u16,
    /// Italic flag
    pub italic: bool,
    /// Underline style (0 = none, 1 = single, 2 = double)
    pub underline: u8,
    /// Font color index
    pub color_index: u16,
    /// Font name
    pub name: String,
}

impl Default for Font {
    fn default() -> Self {
        Self {
            height: 200, // 10pt
            weight: FONT_WEIGHT_NORMAL,
            italic: false,
            underline: 0,
            color_index: COLOR_AUTOMATIC,
            name: "Arial".to_string(),
        }
    }
}

impl Font {
    /// Refuses a font a BIFF8 `Font` record ([MS-XLS] 2.4.122) cannot store
    /// and litchi's reader would refuse: a name that is empty, holds a NUL or
    /// is longer than 31 UTF-16 code units ([`Error::StringTooLong`]); a
    /// height other than 0 or 20 through 8191 twips; a weight other than 0
    /// or 100 through 1000; an underline style other than 0x00, 0x01, 0x02,
    /// 0x21 or 0x22; or a color that is not a font color index.
    pub(crate) fn validate(&self) -> Result<()> {
        if self.name.is_empty() {
            return Err(Error::InvalidData(
                "font name must not be empty".to_string(),
            ));
        }
        ensure_utf16_len_within(&self.name, FONT_NAME_UNITS, "font name")?;
        if self.name.contains('\0') {
            return Err(Error::InvalidData(
                "font name must not contain a NUL character".to_string(),
            ));
        }
        if self.height != 0 && !(20..=8191).contains(&self.height) {
            return Err(Error::InvalidData(format!(
                "font height {} twips is outside 0 and 20 through 8191",
                self.height
            )));
        }
        if self.weight != 0 && !(100..=1000).contains(&self.weight) {
            return Err(Error::InvalidData(format!(
                "font weight {} is outside 0 and 100 through 1000",
                self.weight
            )));
        }
        if !matches!(self.underline, 0x00 | 0x01 | 0x02 | 0x21 | 0x22) {
            return Err(Error::InvalidData(format!(
                "font underline style {:#04x} is not a BIFF8 underline",
                self.underline
            )));
        }
        if !crate::font::valid_color_index(self.color_index) {
            return Err(Error::InvalidData(format!(
                "font color {:#06x} is not a BIFF8 font color index",
                self.color_index
            )));
        }
        Ok(())
    }
}

/// Cell borders
#[derive(Debug, Clone, Default)]
pub struct Borders {
    pub left_style: BorderStyle,
    pub left_color: u16,
    pub right_style: BorderStyle,
    pub right_color: u16,
    pub top_style: BorderStyle,
    pub top_color: u16,
    pub bottom_style: BorderStyle,
    pub bottom_color: u16,
}

/// Cell fill (background)
#[derive(Debug, Clone)]
pub struct Fill {
    pub pattern: FillPattern,
    pub foreground_color: u16,
    pub background_color: u16,
}

impl Default for Fill {
    fn default() -> Self {
        Self {
            pattern: FillPattern::None,
            foreground_color: COLOR_AUTOMATIC,
            background_color: COLOR_AUTOMATIC,
        }
    }
}

/// Extended Format (XF) record - combines font, fill, border, alignment
#[derive(Debug, Clone)]
pub struct ExtendedFormat {
    /// Font index
    pub font_index: u16,
    /// Number format index
    pub format_index: u16,
    /// Horizontal alignment
    pub h_align: HorizontalAlignment,
    /// Vertical alignment
    pub v_align: VerticalAlignment,
    /// Text wrap
    pub text_wrap: bool,
    /// Borders
    pub borders: Borders,
    /// Fill
    pub fill: Fill,
}

impl Default for ExtendedFormat {
    fn default() -> Self {
        Self {
            font_index: 0,
            format_index: 0,
            h_align: HorizontalAlignment::General,
            v_align: VerticalAlignment::Bottom,
            text_wrap: false,
            borders: Borders::default(),
            fill: Fill::default(),
        }
    }
}

/// High-level cell style descriptor used to build reusable styles.
///
/// This is a value-based counterpart to POI's `HSSFCellStyle`: it
/// groups together font, borders, fill, alignment, and an optional
/// number format string. `FormattingManager` converts a `CellStyle`
/// into an `ExtendedFormat` plus FONT and FORMAT records.
#[derive(Debug, Clone)]
pub struct CellStyle {
    /// Font definition used by this style.
    pub font: Font,
    /// Cell borders (styles and colors).
    pub borders: Borders,
    /// Cell fill (background pattern and colors).
    pub fill: Fill,
    /// Horizontal alignment.
    pub h_align: HorizontalAlignment,
    /// Vertical alignment.
    pub v_align: VerticalAlignment,
    /// Whether text is wrapped within the cell.
    pub text_wrap: bool,
    /// Optional number format pattern (e.g. "0.00", "yyyy-mm-dd").
    pub number_format: Option<String>,
}

impl Default for CellStyle {
    fn default() -> Self {
        Self {
            font: Font::default(),
            borders: Borders::default(),
            fill: Fill::default(),
            h_align: HorizontalAlignment::General,
            v_align: VerticalAlignment::Bottom,
            text_wrap: false,
            number_format: None,
        }
    }
}

/// Write FONT record (0x0031)
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `font` - Font definition
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn write_font<W: Write>(writer: &mut W, font: &Font) -> Result<()> {
    // Registration refuses such a font; the encoder refuses it too rather
    // than cutting its name.
    font.validate()?;
    let name = font.name.as_str();
    let name_len = u8_len(utf16_len(name), "font name")?;

    // Fixed payload is 14 bytes of properties:
    // - Height (2) + Attributes (2) + ColorIdx (2) + Weight (2)
    // - Escapement (2) + Underline (1) + Family (1) + Charset (1) + Reserved (1)
    // BIFF8 FONT requires an uncompressed UTF-16LE ShortXLUnicodeString.
    let data_len = record_len("Font", 14 + 1 + 1 + usize::from(name_len) * 2)?;
    super::biff::write_record_header(writer, 0x0031, data_len)?;

    // Font height in twips
    writer.write_all(&font.height.to_le_bytes())?;

    // Option flags (italic, strikeout, etc.)
    let mut flags = 0u16;
    if font.italic {
        flags |= 0x0002;
    }
    writer.write_all(&flags.to_le_bytes())?;

    // Color index
    writer.write_all(&font.color_index.to_le_bytes())?;

    // Font weight
    writer.write_all(&font.weight.to_le_bytes())?;

    // Escapement type (0 = none, 1 = superscript, 2 = subscript)
    writer.write_all(&0u16.to_le_bytes())?;

    // Underline type
    writer.write_all(&[font.underline])?;

    // Font family (0 = None)
    writer.write_all(&[0])?;

    // Character set (0 = ANSI Latin)
    writer.write_all(&[0])?;

    // Reserved
    writer.write_all(&[0])?;

    // Font name length
    writer.write_all(&[name_len])?;

    // FONT narrows ShortXLUnicodeString: fHighByte MUST equal 1.
    writer.write_all(&[0x01])?;

    // Font name UTF-16LE code units.
    for unit in name.encode_utf16() {
        writer.write_all(&unit.to_le_bytes())?;
    }

    Ok(())
}

/// Write XF (Extended Format) record (0x00E0)
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `xf` - Extended format definition
/// * `is_style_xf` - True for style XF, false for cell XF
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn write_xf<W: Write>(writer: &mut W, xf: &ExtendedFormat, is_style_xf: bool) -> Result<()> {
    super::biff::write_record_header(writer, 0x00E0, 20)?;

    // Font index
    writer.write_all(&xf.font_index.to_le_bytes())?;

    // Format index
    writer.write_all(&xf.format_index.to_le_bytes())?;

    // XF type, cell protection, parent style XF
    let xf_type: u16 = if is_style_xf { 0xFFF5 } else { 0x0001 };
    writer.write_all(&xf_type.to_le_bytes())?;

    // Alignment and break
    let mut align_flags = (xf.h_align as u8) | ((xf.v_align as u8) << 4);
    if xf.text_wrap {
        align_flags |= 0x08; // Wrap text bit
    }
    writer.write_all(&[align_flags])?;

    // Rotation
    writer.write_all(&[0])?;

    // Text direction, indent
    writer.write_all(&[0])?;

    // Used attributes flags
    writer.write_all(&[0])?;

    // Border styles bitfield (field_6_border_options).
    // Matches POI's ExtendedFormatRecord: 4-bit nibbles per side.
    let border_left = (xf.borders.left_style as u16) & 0x000F;
    let border_right = ((xf.borders.right_style as u16) & 0x000F) << 4;
    let border_top = ((xf.borders.top_style as u16) & 0x000F) << 8;
    let border_bottom = ((xf.borders.bottom_style as u16) & 0x000F) << 12;
    let border_options = border_left | border_right | border_top | border_bottom;
    writer.write_all(&border_options.to_le_bytes())?;

    // Border palette indices and diagonal flags (field_7_palette_options).
    let left_idx = xf.borders.left_color & 0x007F;
    let right_idx = xf.borders.right_color & 0x007F;
    let palette_options: u16 = (left_idx & 0x007F) | ((right_idx & 0x007F) << 7);
    writer.write_all(&palette_options.to_le_bytes())?;

    // Additional palette options and fill pattern (field_8_adtl_palette_options).
    let top_idx = xf.borders.top_color & 0x007F;
    let bottom_idx = xf.borders.bottom_color & 0x007F;
    let mut adtl_palette_options: u32 = 0;
    adtl_palette_options |= u32::from(top_idx) & 0x0000_007F;
    adtl_palette_options |= (u32::from(bottom_idx) & 0x0000_007F) << 7;
    // Diagonal and diagonal line style are left at 0 (not used).
    let fill_pattern_bits = (xf.fill.pattern as u32) & 0x3F;
    adtl_palette_options |= fill_pattern_bits << 26;
    writer.write_all(&adtl_palette_options.to_le_bytes())?;

    // Fill foreground and background palette indices (field_9_fill_palette_options).
    let fg_idx = xf.fill.foreground_color & 0x007F;
    let bg_idx = xf.fill.background_color & 0x007F;
    let fill_palette_options: u16 = (fg_idx & 0x007F) | ((bg_idx & 0x007F) << 7);
    writer.write_all(&fill_palette_options.to_le_bytes())?;

    Ok(())
}

/// Where a number format pattern would be registered.
enum NumberFormatSlot<'p> {
    /// A built-in or already registered format with this index.
    Existing(u16),
    /// A new custom format with this index and normalized pattern.
    New(u16, &'p str),
}

/// Formatting manager for tracking fonts and formats
#[derive(Debug)]
pub struct FormattingManager {
    fonts: Vec<Font>,
    formats: Vec<ExtendedFormat>,
    // Custom number formats (FORMAT records) keyed by index code.
    // Built-in formats (0x00..0x31) come from BUILTIN_NUMBER_FORMATS.
    number_formats: Vec<(u16, String)>,
    number_format_map: HashMap<String, u16>,
    pivot_xfs_enabled: bool,
}

impl FormattingManager {
    /// Create a new formatting manager with default entries
    #[must_use]
    pub fn new() -> Self {
        let mut manager = Self {
            fonts: Vec::new(),
            formats: Vec::new(),
            number_formats: Vec::new(),
            number_format_map: HashMap::new(),
            pivot_xfs_enabled: false,
        };

        // Add default fonts (indices 0..3) to approximate Excel/POI defaults.
        // 0: Normal
        manager.fonts.push(Font::default());
        // 1: Bold
        manager.fonts.push(Font {
            weight: FONT_WEIGHT_BOLD,
            ..Font::default()
        });
        // 2: Italic
        manager.fonts.push(Font {
            italic: true,
            ..Font::default()
        });
        // 3: Bold + Italic
        manager.fonts.push(Font {
            weight: FONT_WEIGHT_BOLD,
            italic: true,
            ..Font::default()
        });

        // Add default format (index 0)
        manager.formats.push(ExtendedFormat::default());

        manager
    }

    /// Add a font and return its BIFF8 `FontIndex`.
    ///
    /// The first four fonts are the defaults 0 through 3; `FontIndex` skips 4
    /// ([MS-XLS] 2.5.129), so the first added font is 5. Use the returned
    /// value as [`ExtendedFormat::font_index`].
    ///
    /// # Errors
    ///
    /// Refuses a font [`Font`] cannot store in a BIFF8 `Font` record (see its
    /// fields; a name longer than 31 UTF-16 code units is
    /// [`Error::StringTooLong`]) and, with [`Error::TooMany`], a font past
    /// the 1022 a `FontIndex` can address. A refused font leaves the manager
    /// unchanged.
    pub fn add_font(&mut self, font: Font) -> Result<u16> {
        font.validate()?;
        let index = self.next_font_index()?;
        self.fonts.push(font);
        Ok(index)
    }

    /// The `FontIndex` the next added font takes, or [`Error::TooMany`] when
    /// no index is left.
    fn next_font_index(&self) -> Result<u16> {
        if self.fonts.len() >= MAX_FONTS {
            return Err(Error::TooMany {
                collection: "fonts",
                limit: MAX_FONTS,
            });
        }
        let physical = self.fonts.len();
        let logical = if physical < 4 { physical } else { physical + 1 };
        u16::try_from(logical).map_err(|_error| Error::TooMany {
            collection: "fonts",
            limit: MAX_FONTS,
        })
    }

    /// Add a cell format (XF) and return the identifier the `*_with_format`
    /// writer methods take.
    ///
    /// # Errors
    ///
    /// Refuses, with [`Error::InvalidData`], a format whose
    /// [`ExtendedFormat::font_index`] names no font of this manager (4 names
    /// none) or whose [`ExtendedFormat::format_index`] names no built-in or
    /// registered number format, and with [`Error::TooMany`] a format past
    /// the XF records a 16-bit XF index can address. A refused format leaves
    /// the manager unchanged.
    pub fn add_format(&mut self, format: ExtendedFormat) -> Result<u16> {
        self.check_format(&format)?;
        let index = self.next_format_index()?;
        self.formats.push(format);
        Ok(index)
    }

    fn check_format(&self, format: &ExtendedFormat) -> Result<()> {
        if self.get_font(format.font_index).is_none() {
            return Err(Error::InvalidData(format!(
                "cell format refers to font index {}, which names no font",
                format.font_index
            )));
        }
        if !self.contains_number_format_id(format.format_index) {
            return Err(Error::InvalidData(format!(
                "cell format refers to number format {}, which is neither built in nor registered",
                format.format_index
            )));
        }
        Ok(())
    }

    /// The identifier the next added cell format takes, or
    /// [`Error::TooMany`] when the XF index space is full.
    fn next_format_index(&self) -> Result<u16> {
        if self.formats.len() >= MAX_CELL_FORMATS {
            return Err(Error::TooMany {
                collection: "cell formats",
                limit: MAX_CELL_FORMATS - 1,
            });
        }
        u16::try_from(self.formats.len()).map_err(|_error| Error::TooMany {
            collection: "cell formats",
            limit: MAX_CELL_FORMATS - 1,
        })
    }

    /// Register a number format pattern and return its BIFF format index.
    ///
    /// This mirrors POI's `HSSFDataFormat.getFormat` behavior:
    /// - Built-in formats (see `BUILTIN_NUMBER_FORMATS`) return their
    ///   predefined indices.
    /// - The "TEXT" alias normalizes to "@".
    /// - Custom patterns are assigned indices starting at 164 and
    ///   written as FORMAT (0x041E) records.
    ///
    /// # Errors
    ///
    /// A `Format` record holds 1 through 255 UTF-16 code units ([MS-XLS]
    /// 2.4.126): an empty pattern is refused with [`Error::InvalidData`] and a
    /// longer one with [`Error::StringTooLong`]. A workbook holds at most 210
    /// custom formats (the globals grammar allows 218 `Format` records and
    /// the writer always emits eight built-in ones), so a new one past that is
    /// refused with [`Error::TooMany`]. A refused pattern leaves the manager
    /// unchanged.
    pub fn register_number_format(&mut self, pattern: &str) -> Result<u16> {
        match self.number_format_slot(pattern)? {
            NumberFormatSlot::Existing(index) => Ok(index),
            NumberFormatSlot::New(index, normalized) => {
                self.number_formats.push((index, normalized.to_string()));
                self.number_format_map.insert(normalized.to_string(), index);
                Ok(index)
            },
        }
    }

    /// The index `pattern` has or would take, checked as
    /// [`Self::register_number_format`] checks it, without registering it.
    fn number_format_slot<'p>(&self, pattern: &'p str) -> Result<NumberFormatSlot<'p>> {
        validate_number_format(pattern)?;
        // Normalize "TEXT" alias used by POI to "@".
        let normalized = if pattern.eq_ignore_ascii_case("TEXT") {
            "@"
        } else {
            pattern
        };

        // Built-in lookup
        if let Some(idx) = builtin_number_format_index(normalized) {
            return Ok(NumberFormatSlot::Existing(idx));
        }

        // Existing custom format
        if let Some(&idx) = self.number_format_map.get(normalized) {
            return Ok(NumberFormatSlot::Existing(idx));
        }

        if self.number_formats.len() >= MAX_CUSTOM_NUMBER_FORMATS {
            return Err(Error::TooMany {
                collection: "custom number formats",
                limit: MAX_CUSTOM_NUMBER_FORMATS,
            });
        }
        // Allocate new user-defined format index starting at 164, as in BIFF8.
        Ok(NumberFormatSlot::New(
            self.next_custom_format_index(),
            normalized,
        ))
    }

    /// Register a high-level `CellStyle` and return its internal style index.
    ///
    /// This helper wires fonts, number formats, and XF properties together:
    /// - The provided font is appended to the FONT table and its index stored
    ///   in the resulting `ExtendedFormat`.
    /// - If a number format pattern is specified, it is registered via
    ///   `register_number_format` and the resulting index is stored in
    ///   `ExtendedFormat.format_index`.
    /// - Borders, fills, and alignment settings are copied into the XF.
    ///
    /// # Errors
    ///
    /// Refuses the style's number format as [`Self::register_number_format`]
    /// does, its font as [`Self::add_font`] does and its XF as
    /// [`Self::add_format`] does. Every check runs before anything is
    /// registered, so a refused style leaves the manager unchanged.
    pub fn register_cell_style(&mut self, style: CellStyle) -> Result<u16> {
        let CellStyle {
            font,
            borders,
            fill,
            h_align,
            v_align,
            text_wrap,
            number_format,
        } = style;

        font.validate()?;
        let font_index = self.next_font_index()?;
        let xf_index = self.next_format_index()?;
        let format_slot = match number_format.as_deref() {
            Some(pattern) => Some(self.number_format_slot(pattern)?),
            None => None,
        };

        // Every check passed; nothing below can fail.
        let format_index = match format_slot {
            None => 0,
            Some(NumberFormatSlot::Existing(index)) => index,
            Some(NumberFormatSlot::New(index, normalized)) => {
                self.number_formats.push((index, normalized.to_string()));
                self.number_format_map.insert(normalized.to_string(), index);
                index
            },
        };
        self.fonts.push(font);
        self.formats.push(ExtendedFormat {
            font_index,
            format_index,
            h_align,
            v_align,
            text_wrap,
            borders,
            fill,
        });
        Ok(xf_index)
    }

    pub fn enable_pivot_xfs(&mut self) {
        self.pivot_xfs_enabled = true;
    }

    pub(crate) const fn pivot_xfs_enabled(&self) -> bool {
        self.pivot_xfs_enabled
    }

    /// The XF indices of the three pivot-table XFs: index 64 onwards, or the
    /// first index after the user cell XFs when those reach past 63.
    #[must_use]
    pub fn pivot_xf_indices(&self) -> PivotXfIndices {
        let base = xf_record_count_for(self.formats.len(), false).max(PIVOT_XF_START_INDEX);
        // `MAX_CELL_FORMATS` keeps the three pivot XFs inside the index space.
        let base = u16::try_from(base).unwrap_or(u16::MAX - 2);
        PivotXfIndices {
            header_accent: base,
            row_label: base + 1,
            value: base + 2,
        }
    }

    /// Get a font by its BIFF8 `FontIndex` (4 names no font).
    #[must_use]
    pub fn get_font(&self, index: u16) -> Option<&Font> {
        let physical = match index {
            0..=3 => usize::from(index),
            4 => return None,
            _ => usize::from(index) - 1,
        };
        self.fonts.get(physical)
    }

    /// Get format by index
    #[must_use]
    pub fn get_format(&self, index: u16) -> Option<&ExtendedFormat> {
        self.formats.get(index as usize)
    }

    pub(crate) fn contains_number_format_id(&self, index: u16) -> bool {
        BUILTIN_NUMBER_FORMATS
            .get(index as usize)
            .is_some_and(|pattern| !pattern.is_empty())
            || self.number_formats.iter().any(|(id, _)| *id == index)
    }

    /// Write all FONT records
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn write_fonts<W: Write>(&self, writer: &mut W) -> Result<()> {
        for font in &self.fonts {
            write_font(writer, font)?;
        }
        Ok(())
    }

    /// Write all FORMAT records (0x041E): the eight BIFF8 default records and any
    /// registered user-defined formats.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn write_number_formats<W: Write>(&self, writer: &mut W) -> Result<()> {
        // [MS-XLS] 2.4.126 permits the default locale-sensitive groups 5..=8
        // and 41..=44. Built-ins 0..=4 are referenced directly by XF records
        // and MUST NOT be serialized as Format records.
        for index in [5u16, 6, 7, 8, 41, 42, 43, 44] {
            super::biff::write_format_record(
                writer,
                index,
                BUILTIN_NUMBER_FORMATS[index as usize],
            )?;
        }

        for (code, pattern) in &self.number_formats {
            super::biff::write_format_record(writer, *code, pattern)?;
        }

        Ok(())
    }

    /// Write all XF records
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn write_formats<W: Write>(&self, writer: &mut W) -> Result<()> {
        // Base style XF (matches Excel/POI defaults: General format, font 0)
        let base = ExtendedFormat::default();

        // 0..14: default style XFs
        // Map style XFs to fonts in a POI-like way:
        //  0: font 0 (Normal)
        //  1,2: font 1 (Bold)
        //  3,4: font 2 (Italic)
        //  5..14: font 0 (Normal)
        for i in 0..15 {
            let mut xf = base.clone();
            xf.font_index = match i {
                1 | 2 => 1,
                3 | 4 => 2,
                _ => 0,
            };
            write_xf(writer, &xf, true)?;
        }

        // 15: default cell XF used by our cell records
        if let Some(default_cell_xf) = self.formats.first() {
            write_xf(writer, default_cell_xf, false)?;
        } else {
            write_xf(writer, &base, false)?;
        }

        // 16..20: built-in style XFs for common number formats.
        // These mirror POI's mapping of style XFs to built-in number formats,
        // but we keep the structure minimal while remaining BIFF8-compliant.
        //
        // Format indices are BIFF built-ins:
        //  - 0x2B, 0x29, 0x2C, 0x2A: locale-dependent currency/comma styles
        //  - 0x09: percentage
        const BUILTIN_STYLE_FORMATS: [u16; 5] = [0x002B, 0x0029, 0x002C, 0x002A, 0x0009];
        for &fmt_idx in &BUILTIN_STYLE_FORMATS {
            let mut xf = base.clone();
            xf.format_index = fmt_idx;
            write_xf(writer, &xf, true)?;
        }

        // Additional cell XFs (if user-defined formats are ever added later)
        if self.formats.len() > 1 {
            for format in &self.formats[1..] {
                write_xf(writer, format, false)?;
            }
        }

        if self.pivot_xfs_enabled {
            const PIVOT_HEADER_ACCENT: [u8; 20] = [
                0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x02, 0xC0, 0x60,
            ];
            const PIVOT_ROW_LABEL: [u8; 20] = [
                0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x11, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x02, 0xC0, 0x20,
            ];
            const PIVOT_VALUE: [u8; 20] = [
                0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x10, 0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x02, 0xC0, 0x20,
            ];
            // Pad with default cell XFs up to index 64, where the pivot XFs
            // start unless the user cell XFs already reach past it.
            let emitted_xf_count = xf_record_count_for(self.formats.len(), false);
            let pad_count = PIVOT_XF_START_INDEX.saturating_sub(emitted_xf_count);
            let default_payload = if let Some(default_cell_xf) = self.formats.first() {
                let mut buf = Vec::with_capacity(20);
                write_xf(&mut buf, default_cell_xf, false)?;
                let mut payload = [0u8; 20];
                payload.copy_from_slice(&buf[4..24]);
                payload
            } else {
                [
                    0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                    0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0x3F,
                ]
            };

            for _ in 0..pad_count {
                super::biff::write_record_header(writer, 0x00E0, 20)?;
                writer.write_all(&default_payload)?;
            }

            for payload in [PIVOT_HEADER_ACCENT, PIVOT_ROW_LABEL, PIVOT_VALUE] {
                super::biff::write_record_header(writer, 0x00E0, 20)?;
                writer.write_all(&payload)?;
            }
        }

        Ok(())
    }

    /// Total number of XF records `write_formats` emits, including the
    /// pivot padding and extension XFs when pivot XFs are enabled.
    pub(crate) fn xf_record_count(&self) -> usize {
        xf_record_count_for(self.formats.len(), self.pivot_xfs_enabled)
    }

    /// The XF records `write_formats` would emit after adding
    /// `added_formats` cell formats and, when `pivot`, the pivot XFs.
    pub(crate) fn xf_record_count_with(&self, added_formats: usize, pivot: bool) -> usize {
        xf_record_count_for(
            self.formats.len().saturating_add(added_formats),
            pivot || self.pivot_xfs_enabled,
        )
    }

    /// Compute the next available user-defined number format index.
    ///
    /// BIFF8 reserves built-in indices below 164; custom formats start
    /// at `FIRST_USER_DEFINED_NUMBER_FORMAT_INDEX`. We allocate indices
    /// monotonically increasing from that base, mirroring POI's
    /// `InternalWorkbook.getFormat` behavior.
    fn next_custom_format_index(&self) -> u16 {
        self.number_formats
            .iter()
            .map(|(code, _)| *code)
            .max()
            .map_or(FIRST_USER_DEFINED_NUMBER_FORMAT_INDEX, |max_code| {
                max_code.saturating_add(1)
            })
    }
    pub(crate) fn cell_xf_index_for(&self, format_index: u16) -> u16 {
        const STYLE_XF_COUNT: u16 = 15;
        const BUILTIN_STYLE_XF_COUNT: u16 = 5;
        const DEFAULT_CELL_XF_INDEX: u16 = STYLE_XF_COUNT;
        const USER_CELL_XF_START_INDEX: u16 = DEFAULT_CELL_XF_INDEX + 1 + BUILTIN_STYLE_XF_COUNT;

        if format_index == 0 {
            DEFAULT_CELL_XF_INDEX
        } else {
            USER_CELL_XF_START_INDEX + (format_index - 1)
        }
    }
}

impl Default for FormattingManager {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_font_creation() {
        let font = Font {
            name: "Arial".to_string(),
            height: 240, // 12pt
            weight: FONT_WEIGHT_BOLD,
            italic: true,
            ..Default::default()
        };
        assert_eq!(font.name, "Arial");
        assert_eq!(font.height, 240);
        assert!(font.italic);
    }

    #[test]
    fn written_font_round_trips_through_biff8_reader() {
        let font = Font {
            name: "ＭＳ ゴシック".to_string(),
            height: 240,
            weight: FONT_WEIGHT_BOLD,
            italic: true,
            ..Default::default()
        };
        let mut record = Vec::new();
        write_font(&mut record, &font).unwrap();

        let parsed = crate::font::Font::parse_record(
            0,
            &record[4..],
            &mut crate::leniency::ToleranceLog::new(crate::leniency::Leniency::Strict),
        )
        .unwrap();
        assert_eq!(parsed.name(), font.name);
        assert_eq!(parsed.height_twips(), font.height);
        assert!(parsed.is_bold());
        assert!(parsed.is_italic());
    }

    #[test]
    fn test_font_default() {
        let font: Font = Default::default();
        assert_eq!(font.name, "Arial");
        assert_eq!(font.height, 200); // 10pt
        assert_eq!(font.weight, FONT_WEIGHT_NORMAL);
        assert!(!font.italic);
        assert_eq!(font.underline, 0);
        assert_eq!(font.color_index, COLOR_AUTOMATIC);
    }

    #[test]
    fn test_font_clone() {
        let font = Font {
            name: "Arial".to_string(),
            height: 240,
            weight: FONT_WEIGHT_BOLD,
            italic: true,
            ..Default::default()
        };
        let cloned = font.clone();
        assert_eq!(font.name, cloned.name);
        assert_eq!(font.height, cloned.height);
        assert_eq!(font.weight, cloned.weight);
        assert_eq!(font.italic, cloned.italic);
    }

    #[test]
    fn test_border_style_default() {
        let style: BorderStyle = Default::default();
        assert_eq!(style, BorderStyle::None);
    }

    #[test]
    fn test_borders_default() {
        let borders: Borders = Default::default();
        assert_eq!(borders.left_style, BorderStyle::None);
        assert_eq!(borders.right_style, BorderStyle::None);
        assert_eq!(borders.top_style, BorderStyle::None);
        assert_eq!(borders.bottom_style, BorderStyle::None);
    }

    #[test]
    fn test_borders_custom() {
        let borders = Borders {
            left_style: BorderStyle::Thin,
            left_color: COLOR_BLACK,
            right_style: BorderStyle::Medium,
            right_color: COLOR_RED,
            top_style: BorderStyle::Thick,
            top_color: COLOR_WHITE,
            bottom_style: BorderStyle::Double,
            bottom_color: COLOR_AUTOMATIC,
        };
        assert_eq!(borders.left_style, BorderStyle::Thin);
        assert_eq!(borders.right_style, BorderStyle::Medium);
        assert_eq!(borders.top_style, BorderStyle::Thick);
        assert_eq!(borders.bottom_style, BorderStyle::Double);
    }

    #[test]
    fn test_fill_default() {
        let fill: Fill = Default::default();
        assert_eq!(fill.pattern, FillPattern::None);
        assert_eq!(fill.foreground_color, COLOR_AUTOMATIC);
        assert_eq!(fill.background_color, COLOR_AUTOMATIC);
    }

    #[test]
    fn test_fill_solid() {
        let fill = Fill {
            pattern: FillPattern::Solid,
            foreground_color: COLOR_RED,
            background_color: COLOR_WHITE,
        };
        assert_eq!(fill.pattern, FillPattern::Solid);
        assert_eq!(fill.foreground_color, COLOR_RED);
    }

    #[test]
    fn test_extended_format_default() {
        let xf: ExtendedFormat = Default::default();
        assert_eq!(xf.font_index, 0);
        assert_eq!(xf.format_index, 0);
        assert_eq!(xf.h_align, HorizontalAlignment::General);
        assert_eq!(xf.v_align, VerticalAlignment::Bottom);
        assert!(!xf.text_wrap);
    }

    #[test]
    fn test_cell_style_default() {
        let style: CellStyle = Default::default();
        assert_eq!(style.h_align, HorizontalAlignment::General);
        assert_eq!(style.v_align, VerticalAlignment::Bottom);
        assert!(!style.text_wrap);
        assert!(style.number_format.is_none());
    }

    #[test]
    fn test_formatting_manager() {
        let mut mgr = FormattingManager::new();

        let font_idx = mgr
            .add_font(Font {
                name: "Times".to_string(),
                weight: FONT_WEIGHT_BOLD,
                ..Default::default()
            })
            .unwrap();

        // Indices 0..3 are the default fonts, and FontIndex skips 4.
        assert_eq!(font_idx, 5);
        assert_eq!(mgr.get_font(5).unwrap().name, "Times");
        assert!(mgr.get_font(4).is_none());
    }

    #[test]
    fn test_formatting_manager_default_fonts() {
        let mgr = FormattingManager::new();
        // Should have 4 default fonts
        assert_eq!(mgr.get_font(0).unwrap().weight, FONT_WEIGHT_NORMAL);
        assert_eq!(mgr.get_font(1).unwrap().weight, FONT_WEIGHT_BOLD);
        assert!(mgr.get_font(2).unwrap().italic);
        assert!(mgr.get_font(3).unwrap().italic);
        assert_eq!(mgr.get_font(3).unwrap().weight, FONT_WEIGHT_BOLD);
    }

    #[test]
    fn test_formatting_manager_get_font_invalid() {
        let mgr = FormattingManager::new();
        assert!(mgr.get_font(100).is_none());
    }

    #[test]
    fn test_formatting_manager_get_format_invalid() {
        let mgr = FormattingManager::new();
        assert!(mgr.get_format(100).is_none());
    }

    #[test]
    fn test_formatting_manager_add_format() {
        let mut mgr = FormattingManager::new();
        let xf = ExtendedFormat {
            font_index: 1,
            format_index: 2,
            h_align: HorizontalAlignment::Center,
            v_align: VerticalAlignment::Center,
            text_wrap: true,
            ..Default::default()
        };
        let idx = mgr.add_format(xf).unwrap();
        assert_eq!(idx, 1); // Index 0 is default format
        let retrieved = mgr.get_format(1).unwrap();
        assert_eq!(retrieved.font_index, 1);
        assert_eq!(retrieved.h_align, HorizontalAlignment::Center);
        assert!(retrieved.text_wrap);
    }

    #[test]
    fn test_formatting_manager_register_number_format_builtin() {
        let mut mgr = FormattingManager::new();
        // Built-in format should return predefined index
        let idx = mgr.register_number_format("General").unwrap();
        assert_eq!(idx, 0);
        let idx2 = mgr.register_number_format("0.00").unwrap();
        assert_eq!(idx2, 2);
    }

    #[test]
    fn test_formatting_manager_register_number_format_custom() {
        let mut mgr = FormattingManager::new();
        let idx = mgr.register_number_format("0.00\"mm\"").unwrap();
        // Custom formats start at index 164
        assert_eq!(idx, 164);
        // Second registration should return same index
        let idx2 = mgr.register_number_format("0.00\"mm\"").unwrap();
        assert_eq!(idx2, 164);
    }

    #[test]
    fn test_formatting_manager_register_number_format_text_alias() {
        let mut mgr = FormattingManager::new();
        let idx = mgr.register_number_format("TEXT").unwrap();
        assert_eq!(idx, 0x31); // "@" is index 49
    }

    #[test]
    fn a_refused_number_format_leaves_the_manager_unchanged() {
        let mut mgr = FormattingManager::new();
        mgr.register_number_format("0.0\"x\"").unwrap();
        let formats = mgr.number_formats.clone();
        let fonts = mgr.fonts.len();
        let xfs = mgr.formats.len();

        assert!(matches!(
            mgr.register_number_format(""),
            Err(Error::InvalidData(message)) if message.contains("empty")
        ));
        for too_long in [
            "0".repeat(256),
            "é".repeat(256),
            format!("{}😀", "0".repeat(254)),
        ] {
            assert!(matches!(
                mgr.register_number_format(&too_long),
                Err(Error::StringTooLong {
                    field: "number format",
                    utf16_units: 256,
                    limit: 255,
                })
            ));
            let style = CellStyle {
                number_format: Some(too_long.clone()),
                ..Default::default()
            };
            assert!(mgr.register_cell_style(style).is_err());
        }
        assert_eq!(mgr.number_formats, formats);
        assert_eq!(mgr.fonts.len(), fonts);
        assert_eq!(mgr.formats.len(), xfs);
        // 255 units fit, and the next custom index follows the kept one.
        assert_eq!(mgr.register_number_format(&"0".repeat(255)).unwrap(), 165);
    }

    #[test]
    fn the_format_encoder_refuses_what_registration_refuses() {
        // Placed behind registration's own checks.
        let mut empty = FormattingManager::new();
        empty.number_formats.push((164, String::new()));
        assert!(matches!(
            empty.write_number_formats(&mut Vec::new()),
            Err(Error::InvalidData(_))
        ));
        let mut long = FormattingManager::new();
        long.number_formats.push((164, "0".repeat(256)));
        assert!(matches!(
            long.write_number_formats(&mut Vec::new()),
            Err(Error::StringTooLong {
                utf16_units: 256,
                limit: 255,
                ..
            })
        ));
        let mut fits = FormattingManager::new();
        fits.number_formats.push((164, "0".repeat(255)));
        fits.write_number_formats(&mut Vec::new()).unwrap();
    }

    #[test]
    fn test_formatting_manager_register_cell_style() {
        let mut mgr = FormattingManager::new();
        let style = CellStyle {
            font: Font {
                name: "Courier".to_string(),
                height: 220,
                ..Default::default()
            },
            h_align: HorizontalAlignment::Right,
            v_align: VerticalAlignment::Top,
            number_format: Some("0.00".to_string()),
            ..Default::default()
        };
        let idx = mgr.register_cell_style(style).unwrap();
        assert_eq!(idx, 1); // Index 0 is default
        let retrieved = mgr.get_format(1).unwrap();
        assert_eq!(retrieved.h_align, HorizontalAlignment::Right);
        assert_eq!(retrieved.v_align, VerticalAlignment::Top);
        assert_eq!(retrieved.format_index, 2); // "0.00" is built-in index 2
    }

    #[test]
    fn test_formatting_manager_pivot_xfs() {
        let mut mgr = FormattingManager::new();
        assert!(!mgr.pivot_xfs_enabled);
        mgr.enable_pivot_xfs();
        let indices = mgr.pivot_xf_indices();
        assert_eq!(indices.header_accent, 64);
        assert_eq!(indices.row_label, 65);
        assert_eq!(indices.value, 66);
    }

    #[test]
    fn test_builtin_number_format_index() {
        assert_eq!(builtin_number_format_index("General"), Some(0));
        assert_eq!(builtin_number_format_index("0"), Some(1));
        assert_eq!(builtin_number_format_index("0.00"), Some(2));
        assert_eq!(builtin_number_format_index("@"), Some(0x31));
        assert_eq!(builtin_number_format_index("NonExistent"), None);
    }

    #[test]
    fn test_horizontal_alignment_variants() {
        assert_eq!(HorizontalAlignment::General as u8, 0);
        assert_eq!(HorizontalAlignment::Left as u8, 1);
        assert_eq!(HorizontalAlignment::Center as u8, 2);
        assert_eq!(HorizontalAlignment::Right as u8, 3);
        assert_eq!(HorizontalAlignment::Fill as u8, 4);
        assert_eq!(HorizontalAlignment::Justify as u8, 5);
        assert_eq!(HorizontalAlignment::CenterAcrossSelection as u8, 6);
    }

    #[test]
    fn test_vertical_alignment_variants() {
        assert_eq!(VerticalAlignment::Top as u8, 0);
        assert_eq!(VerticalAlignment::Center as u8, 1);
        assert_eq!(VerticalAlignment::Bottom as u8, 2);
        assert_eq!(VerticalAlignment::Justify as u8, 3);
    }

    #[test]
    fn test_fill_pattern_variants() {
        assert_eq!(FillPattern::None as u8, 0);
        assert_eq!(FillPattern::Solid as u8, 1);
        assert_eq!(FillPattern::MediumGray as u8, 2);
        assert_eq!(FillPattern::DarkGray as u8, 3);
        assert_eq!(FillPattern::LightGray as u8, 4);
    }

    #[test]
    fn test_border_style_variants() {
        assert_eq!(BorderStyle::None as u8, 0);
        assert_eq!(BorderStyle::Thin as u8, 1);
        assert_eq!(BorderStyle::Medium as u8, 2);
        assert_eq!(BorderStyle::Dashed as u8, 3);
        assert_eq!(BorderStyle::Dotted as u8, 4);
        assert_eq!(BorderStyle::Thick as u8, 5);
        assert_eq!(BorderStyle::Double as u8, 6);
        assert_eq!(BorderStyle::Hair as u8, 7);
    }

    #[test]
    fn test_formatting_manager_default() {
        let mgr: FormattingManager = Default::default();
        assert!(mgr.get_font(0).is_some());
    }

    #[test]
    fn test_write_font_basic() {
        let mut buf = Vec::new();
        let font = Font::default();
        write_font(&mut buf, &font).unwrap();

        // Check record type (0x0031 = FONT)
        assert_eq!(u16::from_le_bytes([buf[0], buf[1]]), 0x0031);

        // Check height
        let height = u16::from_le_bytes([buf[4], buf[5]]);
        assert_eq!(height, 200); // 10pt = 200 twips
    }

    #[test]
    fn test_write_font_bold() {
        let mut buf = Vec::new();
        let font = Font {
            weight: FONT_WEIGHT_BOLD,
            ..Default::default()
        };
        write_font(&mut buf, &font).unwrap();

        let weight = u16::from_le_bytes([buf[10], buf[11]]);
        assert_eq!(weight, FONT_WEIGHT_BOLD);
    }

    #[test]
    fn test_write_font_italic() {
        let mut buf = Vec::new();
        let font = Font {
            italic: true,
            ..Default::default()
        };
        write_font(&mut buf, &font).unwrap();

        let flags = u16::from_le_bytes([buf[6], buf[7]]);
        assert!(flags & 0x0002 != 0); // Italic flag
    }

    #[test]
    fn test_write_xf() {
        let mut buf = Vec::new();
        let xf = ExtendedFormat::default();
        write_xf(&mut buf, &xf, false).unwrap();

        // Check record type (0x00E0 = XF)
        assert_eq!(u16::from_le_bytes([buf[0], buf[1]]), 0x00E0);
        assert_eq!(u16::from_le_bytes([buf[2], buf[3]]), 20); // Length = 20
    }

    #[test]
    fn test_write_xf_style() {
        let mut buf = Vec::new();
        let xf = ExtendedFormat::default();
        write_xf(&mut buf, &xf, true).unwrap();

        // For style XF, xf_type should be 0xFFF5
        let xf_type = u16::from_le_bytes([buf[8], buf[9]]);
        assert_eq!(xf_type, 0xFFF5);
    }

    #[test]
    fn test_write_xf_with_alignment() {
        let mut buf = Vec::new();
        let xf = ExtendedFormat {
            h_align: HorizontalAlignment::Center,
            v_align: VerticalAlignment::Center,
            text_wrap: true,
            ..Default::default()
        };
        write_xf(&mut buf, &xf, false).unwrap();

        // Check alignment byte
        let align = buf[10];
        assert_eq!(align & 0x07, 2); // Horizontal center
        assert!(align & 0x08 != 0); // Text wrap
    }

    #[test]
    fn test_cell_xf_index_for() {
        let mgr = FormattingManager::new();
        // Default format (index 0) maps to cell XF index 15
        assert_eq!(mgr.cell_xf_index_for(0), 15);
        // Custom format index 1 maps to index 21 (15 + 1 + 5)
        assert_eq!(mgr.cell_xf_index_for(1), 21);
    }

    #[test]
    fn test_pivot_xf_indices_default() {
        let mgr = FormattingManager::new();
        let indices = mgr.pivot_xf_indices();
        assert_eq!(indices.header_accent, 64);
        assert_eq!(indices.row_label, 65);
        assert_eq!(indices.value, 66);
    }
}
