//! MTEF constants and tag definitions
//!
//! This module defines constants used in MathType Equation Format (MTEF) binary parsing.
//! Based on rtf2latex2e implementation and MTEF specification.
//!
//! References:
//! - http://rtf2latex2e.sourceforge.net/MTEF5.html
//! - rtf2latex2e source code (eqn_support.h)

// This module documents the format, not just the subset the current reader and
// writer happen to touch: neighbouring tags, flags and selectors are kept so the
// tables stay readable and reviewable against the specification.
#![allow(dead_code)]

/// MTEF record tags - these identify different types of equation objects
pub const END: u8 = 0;
pub const LINE: u8 = 1;
pub const CHAR: u8 = 2;
pub const TMPL: u8 = 3;
pub const PILE: u8 = 4;
pub const MATRIX: u8 = 5;
pub const EMBELL: u8 = 6;
pub const RULER: u8 = 7;
pub const SIZE: u8 = 9;
pub const FULL: u8 = 10;
pub const SUB: u8 = 11;
pub const SUB2: u8 = 12;
pub const SYM: u8 = 13;
pub const SUBSYM: u8 = 14;
pub const COLOR: u8 = 15;
pub const COLOR_DEF: u8 = 16;
pub const FONT_DEF: u8 = 17;
pub const FONT: u8 = 18;
pub const EQN_PREFS: u8 = 19;
pub const ENCODING_DEF: u8 = 20;

/// Character attribute flags for MTEF character records
pub const CHAR_EMBELL: u8 = 0x01;
pub const CHAR_ENC_CHAR_8: u8 = 0x04;
pub const CHAR_NUDGE: u8 = 0x08;
pub const CHAR_ENC_CHAR_16: u8 = 0x10;
pub const CHAR_ENC_NO_MTCODE: u8 = 0x20;

/// General attribute flags (xf prefix)
pub const XF_EMBELL: u8 = 0x01;
pub const XF_RULER: u8 = 0x02;
pub const XF_NULL: u8 = 0x04;
pub const XF_LSPACE: u8 = 0x04;
pub const XF_LMOVE: u8 = 0x08;

/// Math attribute constants for character set handling
pub const MA_TEXT: i32 = 0; // Text mode
pub const MA_MATH: i32 = 1; // Math mode
pub const MA_FORCE_TEXT: i32 = 2; // Force text mode
pub const MA_FORCE_MATH: i32 = 3; // Force math mode

/// Number of typeface slots in MTEF
pub const NUM_TYPEFACE_SLOTS: usize = 32;

/// Equation mode constants for mode switching during parsing
pub const EQN_MODE_TEXT: i32 = 0;
pub const EQN_MODE_INLINE: i32 = 1;
pub const EQN_MODE_DISPLAY: i32 = 2;

// ---------------------------------------------------------------------------
// Record framing
// ---------------------------------------------------------------------------

/// MTEF version that carries the record options in a byte of its own
///
/// Up to MTEF 4 the record tag byte holds the record type in its low nibble and
/// the record attributes in its high nibble. From MTEF 5 on, the tag byte holds
/// the record type alone and is followed by a dedicated options byte.
pub const MTEF_VERSION_5: u8 = 5;

/// Mask selecting the record type from a pre-MTEF-5 tag byte
pub const TAG_TYPE_MASK: u8 = 0x0F;
/// Mask selecting the record attributes from a pre-MTEF-5 tag byte
pub const TAG_ATTRIBUTE_MASK: u8 = 0xF0;
/// Bit shift applied to the attribute nibble of a pre-MTEF-5 tag byte
pub const TAG_ATTRIBUTE_SHIFT: u32 = 4;

/// Sentinel written for both nudge bytes when the extended 16-bit form follows
pub const NUDGE_ESCAPE: u8 = 128;

/// Flag marking a template variation byte as the low half of a 15-bit value
pub const VARIATION_EXTENDED_FLAG: u8 = 0x80;
/// Mask selecting the low bits carried by the first template variation byte
pub const VARIATION_LOW_MASK: u16 = 0x7F;
/// Number of variation bits carried by the first template variation byte
pub const VARIATION_LOW_BITS: u32 = 7;
/// Largest template variation representable by the two-byte encoding
pub const VARIATION_MAX: u16 = 0x7FFF;

/// Bytes available for a MATRIX row or column partition
pub const MATRIX_PARTITION_BYTES: usize = 16;
/// Bits used by a single MATRIX row or column partition entry
pub const MATRIX_PARTITION_BITS: usize = 2;

// ---------------------------------------------------------------------------
// OLE ("Equation Native") stream header
// ---------------------------------------------------------------------------

/// Length of the OLE equation header that precedes MTEF data
pub const OLE_HEADER_LEN: usize = 28;
/// Value stored in the OLE header `cb_hdr` field
pub const OLE_HEADER_CB_HDR: u16 = 28;
/// Value stored in the OLE header `version` field (hiword = 2, loword = 0)
pub const OLE_HEADER_VERSION: u32 = 0x0002_0000;
/// Clipboard format registered for "MathType EF"
pub const OLE_CLIPBOARD_FORMAT: u16 = 0xC1C6;

// ---------------------------------------------------------------------------
// MTEF header
// ---------------------------------------------------------------------------

/// Platform identifier for Macintosh-authored equations
pub const PLATFORM_MACINTOSH: u8 = 0;
/// Platform identifier for Windows-authored equations
pub const PLATFORM_WINDOWS: u8 = 1;
/// Product identifier for MathType (as opposed to the Equation Editor)
pub const PRODUCT_MATHTYPE: u8 = 1;
/// Product major version recorded in the MTEF header
pub const PRODUCT_VERSION: u8 = 5;
/// Product minor version recorded in the MTEF header
pub const PRODUCT_SUB_VERSION: u8 = 0;
/// Header flag marking an equation as display (rather than inline) material
pub const EQUATION_DISPLAY: u8 = 0;
/// Header flag marking an equation as inline material
pub const EQUATION_INLINE: u8 = 1;

// ---------------------------------------------------------------------------
// Typeface slots (character record `typeface` byte = 128 + slot)
// ---------------------------------------------------------------------------

/// Value added to a typeface slot number to form the record's `typeface` byte.
///
/// Slot 0 is the unused ZERO slot, so `TYPEFACE_TEXT` (slot 1) is 129.
pub const TYPEFACE_SLOT_BASE: u8 = 128;

/// Typeface slot for plain text runs
pub const TYPEFACE_TEXT: u8 = 129;
/// Typeface slot for recognised function names (`sin`, `log`, ...)
pub const TYPEFACE_FUNCTION: u8 = 130;
/// Typeface slot for ordinary math variables
pub const TYPEFACE_VARIABLE: u8 = 131;
/// Typeface slot for lowercase Greek letters
pub const TYPEFACE_LCGREEK: u8 = 132;
/// Typeface slot for uppercase Greek letters
pub const TYPEFACE_UCGREEK: u8 = 133;
/// Typeface slot for the Symbol font (operators and relations)
pub const TYPEFACE_SYMBOL: u8 = 134;
/// Typeface slot for bold (vector) characters
pub const TYPEFACE_VECTOR: u8 = 135;
/// Typeface slot for digits
pub const TYPEFACE_NUMBER: u8 = 136;
/// Typeface slot for the MT Extra font (extended operators and dots)
pub const TYPEFACE_MTEXTRA: u8 = 139;
/// Typeface slot for fixed-width spaces
pub const TYPEFACE_SPACE: u8 = 152;

// ---------------------------------------------------------------------------
// MTEF 5 template selectors
// ---------------------------------------------------------------------------
//
// MTEF 5 renumbered the template table that MTEF 1-4 used, so these values are
// only valid for `mtef_version >= MTEF_VERSION_5`.

/// Angle-bracket fence template
pub const TMPL_ANGLE: u8 = 0;
/// Parenthesis fence template
pub const TMPL_PAREN: u8 = 1;
/// Curly-brace fence template
pub const TMPL_BRACE: u8 = 2;
/// Square-bracket fence template
pub const TMPL_BRACK: u8 = 3;
/// Single vertical bar fence template
pub const TMPL_BAR: u8 = 4;
/// Double vertical bar fence template
pub const TMPL_DBAR: u8 = 5;
/// Floor fence template
pub const TMPL_FLOOR: u8 = 6;
/// Ceiling fence template
pub const TMPL_CEILING: u8 = 7;
/// Open (white) bracket fence template
pub const TMPL_OBRACK: u8 = 8;
/// Mixed-delimiter interval fence template
pub const TMPL_INTERVAL: u8 = 9;
/// Square and nth root template
pub const TMPL_ROOT: u8 = 10;
/// Fraction template
pub const TMPL_FRACT: u8 = 11;
/// Underbar template
pub const TMPL_UBAR: u8 = 12;
/// Overbar template
pub const TMPL_OBAR: u8 = 13;
/// Arrow-over-expression template
pub const TMPL_ARROW: u8 = 14;
/// Integral family template (single, double, triple, contour, ...)
pub const TMPL_INTEG: u8 = 15;
/// Summation template with limits above and below
pub const TMPL_SUM: u8 = 16;
/// Product template with limits above and below
pub const TMPL_PROD: u8 = 17;
/// Coproduct template with limits above and below
pub const TMPL_COPROD: u8 = 18;
/// Union template with limits above and below
pub const TMPL_UNION: u8 = 19;
/// Intersection template with limits above and below
pub const TMPL_INTER: u8 = 20;
/// Integral template with limits beside the operator
pub const TMPL_INTOP: u8 = 21;
/// Summation template with limits beside the operator
pub const TMPL_SUMOP: u8 = 22;
/// Limit template (base with material below and/or above it)
pub const TMPL_LIM: u8 = 23;
/// Horizontal brace template
pub const TMPL_HBRACE: u8 = 24;
/// Horizontal bracket template
pub const TMPL_HBRACK: u8 = 25;
/// Subscript template attached to the preceding object
pub const TMPL_SUB: u8 = 27;
/// Superscript template attached to the preceding object
pub const TMPL_SUP: u8 = 28;
/// Combined subscript/superscript template attached to the preceding object
pub const TMPL_SUBSUP: u8 = 29;

// ---------------------------------------------------------------------------
// MTEF 5 template variations
// ---------------------------------------------------------------------------

/// Fence variation: only the opening delimiter is drawn
pub const TV_FENCE_LEFT: u16 = 1;
/// Fence variation: only the closing delimiter is drawn
pub const TV_FENCE_RIGHT: u16 = 2;
/// Fence variation: both delimiters are drawn
pub const TV_FENCE_BOTH: u16 = 3;

/// Interval variation nibble for a left parenthesis
pub const TV_INTERVAL_LPAREN: u16 = 0;
/// Interval variation nibble for a right parenthesis
pub const TV_INTERVAL_RPAREN: u16 = 1;
/// Interval variation nibble for a left square bracket
pub const TV_INTERVAL_LBRACK: u16 = 2;
/// Interval variation nibble for a right square bracket
pub const TV_INTERVAL_RBRACK: u16 = 3;
/// Bit shift applied to the closing-delimiter nibble of an interval variation
pub const TV_INTERVAL_CLOSE_SHIFT: u32 = 4;

/// Root variation: square root (one slot)
pub const TV_ROOT_SQUARE: u16 = 0;
/// Root variation: nth root (radicand and degree slots)
pub const TV_ROOT_NTH: u16 = 1;

/// Fraction variation: full-size fraction with a horizontal bar
pub const TV_FRACT_BAR: u16 = 0;
/// Fraction variation: reduced-size fraction with a horizontal bar
pub const TV_FRACT_SMALL_BAR: u16 = 1;
/// Fraction variation: slashed (linear) fraction without a bar
pub const TV_FRACT_SLASH: u16 = 2;

/// Bar variation: a single rule
pub const TV_BAR_SINGLE: u16 = 0;
/// Bar variation: a double rule
pub const TV_BAR_DOUBLE: u16 = 1;

/// Integral variation: one integral sign, limits suppressed
pub const TV_INTEG_SINGLE: u16 = 0;
/// Integral variation: one integral sign with limits
pub const TV_INTEG_SINGLE_LIMITS: u16 = 1;
/// Integral variation: two integral signs
pub const TV_INTEG_DOUBLE: u16 = 2;
/// Integral variation: three integral signs
pub const TV_INTEG_TRIPLE: u16 = 3;
/// Integral variation: contour integral
pub const TV_INTEG_CONTOUR: u16 = 4;
/// Integral variation: surface integral
pub const TV_INTEG_SURFACE: u16 = 8;
/// Integral variation: volume integral
pub const TV_INTEG_VOLUME: u16 = 12;

/// Horizontal brace variation: brace below the expression
pub const TV_HBRACE_LOWER: u16 = 0;
/// Horizontal brace variation: brace above the expression
pub const TV_HBRACE_UPPER: u16 = 1;

/// Variation used by templates that define exactly one form
pub const TV_DEFAULT: u16 = 0;

// ---------------------------------------------------------------------------
// Embellishments
// ---------------------------------------------------------------------------

/// Single dot above (`\dot`)
pub const EMB_DOT: u8 = 2;
/// Two dots above (`\ddot`)
pub const EMB_DDOT: u8 = 3;
/// Three dots above (`\dddot`)
pub const EMB_TDOT: u8 = 4;
/// Single prime
pub const EMB_PRIME: u8 = 5;
/// Reversed prime
pub const EMB_BPRIME: u8 = 7;
/// Tilde above (`\tilde`)
pub const EMB_TILDE: u8 = 8;
/// Circumflex above (`\hat`)
pub const EMB_HAT: u8 = 9;
/// Right arrow above (`\vec`)
pub const EMB_RARROW: u8 = 11;
/// Bar above (`\bar`)
pub const EMB_OBAR: u8 = 17;
/// Breve above (`\breve`)
pub const EMB_SMILE: u8 = 20;

// ---------------------------------------------------------------------------
// Alignment
// ---------------------------------------------------------------------------

/// Horizontal alignment: centred
pub const ALIGN_CENTER: u8 = 2;
/// Vertical alignment: baseline of the first line
pub const VALIGN_BASELINE: u8 = 1;
/// Justification: centred within the cell
pub const JUSTIFY_CENTER: u8 = 0;
