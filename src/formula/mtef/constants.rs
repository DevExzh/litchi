//! MTEF constants and tag definitions
//!
//! This module defines constants used in MathType Equation Format (MTEF) binary parsing.
//! Based on rtf2latex2e implementation and MTEF specification.
//!
//! References:
//! - http://rtf2latex2e.sourceforge.net/MTEF5.html
//! - rtf2latex2e source code (eqn_support.h)

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
pub const MA_TEXT: i32 = 0;        // Text mode
pub const MA_MATH: i32 = 1;        // Math mode
pub const MA_FORCE_TEXT: i32 = 2;  // Force text mode
pub const MA_FORCE_MATH: i32 = 3;  // Force math mode

/// Number of typeface slots in MTEF
pub const NUM_TYPEFACE_SLOTS: usize = 32;

/// Equation mode constants for mode switching during parsing
pub const EQN_MODE_TEXT: i32 = 0;
pub const EQN_MODE_INLINE: i32 = 1;
pub const EQN_MODE_DISPLAY: i32 = 2;

