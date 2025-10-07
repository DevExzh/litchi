// MTEF constants and tag definitions

// MTEF record tags
pub const END: u8 = 0;
pub const LINE: u8 = 1;
pub const CHAR: u8 = 2;
pub const TMPL: u8 = 3;
pub const PILE: u8 = 4;
pub const MATRIX: u8 = 5;
pub const EMBELL: u8 = 6;
pub const COLOR_DEF: u8 = 16;
pub const FONT: u8 = 18;
pub const FONT_DEF: u8 = 17;
pub const EQN_PREFS: u8 = 19;
pub const ENCODING_DEF: u8 = 20;
// NOTE: Some constants below are currently unused but kept for future MTEF functionality
#[allow(dead_code)]
pub const FUTURE: u8 = 255; // Future record type - kept for future use
pub const RULER: u8 = 7;
pub const SIZE: u8 = 9;
pub const FULL: u8 = 10;
pub const SUB: u8 = 11;
pub const SUB2: u8 = 12;
pub const SYM: u8 = 13;
pub const SUBSYM: u8 = 14;
pub const COLOR: u8 = 15;

// Character attributes
pub const CHAR_EMBELL: u8 = 0x01;
#[allow(dead_code)]
pub const CHAR_FUNC_START: u8 = 0x02; // Function start character attribute - kept for future use
pub const CHAR_ENC_CHAR_8: u8 = 0x04;
pub const CHAR_NUDGE: u8 = 0x08;
pub const CHAR_ENC_CHAR_16: u8 = 0x10;
pub const CHAR_ENC_NO_MTCODE: u8 = 0x20;

// General attributes (xf prefix for attributes)
pub const XF_LMOVE: u8 = 0x08;
pub const XF_EMBELL: u8 = 0x01;
pub const XF_LSPACE: u8 = 0x04;
pub const XF_RULER: u8 = 0x02;
pub const XF_NULL: u8 = 0x04;

// Template selectors
// NOTE: Template constants below are currently unused but kept for future MTEF functionality
#[allow(dead_code)]
pub const TMPL_FRAC: u8 = 0; // Fraction template - kept for future use
#[allow(dead_code)]
pub const TMPL_OVER: u8 = 1; // Over template - kept for future use
#[allow(dead_code)]
pub const TMPL_SLASH: u8 = 2; // Slash template - kept for future use
#[allow(dead_code)]
pub const TMPL_ROOT: u8 = 3; // Root template - kept for future use
#[allow(dead_code)]
pub const TMPL_SUB: u8 = 4; // Subscript template - kept for future use
#[allow(dead_code)]
pub const TMPL_SUP: u8 = 5; // Superscript template - kept for future use
#[allow(dead_code)]
pub const TMPL_SUBSUP: u8 = 6; // Subscript-superscript template - kept for future use
#[allow(dead_code)]
pub const TMPL_SUPSUB: u8 = 7; // Superscript-subscript template - kept for future use
#[allow(dead_code)]
pub const TMPL_BELOW: u8 = 8; // Below template - kept for future use
#[allow(dead_code)]
pub const TMPL_ABOVE: u8 = 9; // Above template - kept for future use
#[allow(dead_code)]
pub const TMPL_BELOABOVE: u8 = 10; // Below-above template - kept for future use
// NOTE: Template constants below are currently unused but kept for future MTEF functionality
// These represent various mathematical template types that may be needed for complete MTEF parsing
#[allow(dead_code)]
pub const TMPL_LPAREN: u8 = 11; // Left parenthesis template - kept for future use
#[allow(dead_code)]
pub const TMPL_RPAREN: u8 = 12; // Right parenthesis template - kept for future use
#[allow(dead_code)]
pub const TMPL_OVERBRACE: u8 = 13; // Overbrace template - kept for future use
#[allow(dead_code)]
pub const TMPL_UNDERBRACE: u8 = 14; // Underbrace template - kept for future use
// Template selectors based on rtf2latex2e Profile_TEMPLATES_5 (these are actually used)
pub const TMPL_LIM: u8 = 23; // Limit template
pub const TMPL_ARROW: u8 = 14; // Arrow template
pub const TMPL_PAREN: u8 = 1; // Parenthesis fence template
pub const TMPL_BRACKET: u8 = 3; // Bracket fence template
pub const TMPL_BRACE: u8 = 2; // Brace fence template
pub const TMPL_BAR: u8 = 4; // Bar fence template
pub const TMPL_DBAR: u8 = 5; // Double bar fence template

// Large operator selectors
pub const TMPL_SUM: u8 = 16; // Summation template
pub const TMPL_PROD: u8 = 17; // Product template
pub const TMPL_COPROD: u8 = 18; // Coproduct template
pub const TMPL_UNION: u8 = 19; // Union template
pub const TMPL_INTER: u8 = 20; // Intersection template
pub const TMPL_INTOP: u8 = 15; // Integral operator template (single, double, triple, contour)
pub const TMPL_IINTOP: u8 = 21; // Double integral (single with limits)
pub const TMPL_IIINTOP: u8 = 22; // Triple integral (single sum with limits)
pub const TMPL_OINTOP: u8 = 23; // Contour integral (limit template)

// Math attribute constants (for character set handling)
pub const MA_TEXT: i32 = 0;        // Text mode
pub const MA_MATH: i32 = 1;        // Math mode
pub const MA_FORCE_TEXT: i32 = 2;  // Force text mode
pub const MA_FORCE_MATH: i32 = 3;  // Force math mode

// Embellishment types (index into EMBELLISHMENT_TEMPLATES)
// NOTE: Embellishment constants below are currently unused but kept for future MTEF functionality
// These represent various mathematical embellishment types that may be needed for complete MTEF parsing
#[allow(dead_code)]
pub const EMBELL_DOT: u8 = 2; // Dot embellishment - kept for future use
#[allow(dead_code)]
pub const EMBELL_DDOT: u8 = 3; // Double dot embellishment - kept for future use
#[allow(dead_code)]
pub const EMBELL_TDOT: u8 = 4; // Triple dot embellishment - kept for future use
#[allow(dead_code)]
pub const EMBELL_PRIME: u8 = 5; // Prime embellishment - kept for future use
#[allow(dead_code)]
pub const EMBELL_DPRIME: u8 = 6; // Double prime embellishment - kept for future use
#[allow(dead_code)]
pub const EMBELL_BPRIME: u8 = 7; // Back prime embellishment - kept for future use
#[allow(dead_code)]
pub const EMBELL_TILDE: u8 = 8; // Tilde embellishment - kept for future use
#[allow(dead_code)]
pub const EMBELL_HAT: u8 = 9; // Hat embellishment - kept for future use
#[allow(dead_code)]
pub const EMBELL_VEC: u8 = 11; // Vector embellishment - kept for future use
#[allow(dead_code)]
pub const EMBELL_BAR: u8 = 16; // Bar embellishment - kept for future use

// Number of typeface slots
pub const NUM_TYPEFACE_SLOTS: usize = 32;

// Equation mode constants (for mode switching)
pub const EQN_MODE_TEXT: i32 = 0;
pub const EQN_MODE_INLINE: i32 = 1;
pub const EQN_MODE_DISPLAY: i32 = 2;

