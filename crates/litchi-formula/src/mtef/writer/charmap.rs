//! AST symbol/operator/character to MTEF typeface and character code mapping
//!
//! This is the inverse of [`crate::mtef::binary::charset`]: that module maps a
//! `(typeface, character)` pair to LaTeX, this one picks the pair that a reader
//! will resolve back to the intended symbol.
//!
//! Two encodings coexist in MathType data and in the reader's tables: the
//! historical Symbol/MT-Extra font positions (`134.163` is `\leq`) and Unicode
//! MTCode values (`134.8804` is also `\leq`). Unicode is preferred here because
//! it survives font substitution; the font positions are used only where the
//! reader has no Unicode entry.

use super::error::MtefWriteError;
use crate::ast::{AccentType, FunctionName, Operator, PredefinedSymbol, SpaceType, StyleType};
use crate::mtef::constants::*;

/// A character as stored by a CHAR record
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct CharCode {
    /// Typeface slot the character is drawn from
    pub typeface: u8,
    /// 16-bit MathType character code
    pub mtcode: u16,
}

impl CharCode {
    /// Build a character code from a typeface slot and MTCode
    const fn new(typeface: u8, mtcode: u16) -> Self {
        Self { typeface, mtcode }
    }
}

/// First lowercase Greek letter in Unicode (α)
const GREEK_LOWER_FIRST: u32 = 0x03B1;
/// Last lowercase Greek letter in Unicode (ω)
const GREEK_LOWER_LAST: u32 = 0x03C9;
/// First uppercase Greek letter in Unicode (Α)
const GREEK_UPPER_FIRST: u32 = 0x0391;
/// Last uppercase Greek letter in Unicode (Ω)
const GREEK_UPPER_LAST: u32 = 0x03A9;

/// Map an arbitrary character to a typeface slot and MTCode
///
/// Digits go to the NUMBER slot and other ASCII to the VARIABLE slot, which the
/// reader renders verbatim. Greek letters use the dedicated Greek slots so that
/// they come back as `\alpha`-style commands, and everything else is offered to
/// the Symbol slot, whose lookup table covers the mathematical operators and
/// falls back to the literal codepoint.
pub(super) fn char_code(ch: char) -> Result<CharCode, MtefWriteError> {
    let code = u32::from(ch);
    let mtcode = u16::try_from(code).map_err(|_| MtefWriteError::UnsupportedCharacter(ch))?;

    let typeface = match code {
        _ if ch.is_ascii_digit() => TYPEFACE_NUMBER,
        _ if ch.is_ascii() => TYPEFACE_VARIABLE,
        GREEK_LOWER_FIRST..=GREEK_LOWER_LAST => TYPEFACE_LCGREEK,
        GREEK_UPPER_FIRST..=GREEK_UPPER_LAST => TYPEFACE_UCGREEK,
        _ => TYPEFACE_SYMBOL,
    };

    Ok(CharCode::new(typeface, mtcode))
}

/// Map a character that belongs to a run with an explicit typeface
///
/// Style runs pin the typeface (bold text is the VECTOR slot, for instance) but
/// still need the MTCode, and characters outside the BMP remain unencodable.
pub(super) fn char_code_in(ch: char, typeface: u8) -> Result<CharCode, MtefWriteError> {
    let mtcode =
        u16::try_from(u32::from(ch)).map_err(|_| MtefWriteError::UnsupportedCharacter(ch))?;
    Ok(CharCode::new(typeface, mtcode))
}

/// Typeface slot that renders a style, if MTEF can express it
///
/// MTEF encodes emphasis through the typeface slot rather than through record
/// attributes, so only the styles that own a slot survive; the rest fall back to
/// the default variable slot.
pub(super) fn style_typeface(style: StyleType) -> Option<u8> {
    match style {
        StyleType::Bold | StyleType::BoldItalic => Some(TYPEFACE_VECTOR),
        StyleType::Normal => Some(TYPEFACE_TEXT),
        _ => None,
    }
}

/// Map an operator to its MTEF character
///
/// Codes taken from the reader's lookup table; where both a Unicode and a
/// Symbol-font entry exist the Unicode one is preferred.
pub(super) fn operator_code(operator: Operator) -> CharCode {
    /// Shorthand for a Symbol-slot character
    const fn symbol(mtcode: u16) -> CharCode {
        CharCode::new(TYPEFACE_SYMBOL, mtcode)
    }
    /// Shorthand for an MT-Extra-slot character
    const fn extra(mtcode: u16) -> CharCode {
        CharCode::new(TYPEFACE_MTEXTRA, mtcode)
    }

    match operator {
        // Arithmetic
        Operator::Plus => symbol(0x002B),
        Operator::Minus => symbol(0x2212),
        Operator::Multiply | Operator::Times | Operator::Cross => symbol(180), // Symbol font "\times"
        Operator::Divide => symbol(184),                                       // Symbol font "\div"
        Operator::PlusMinus => symbol(177),                                    // Symbol font "\pm"
        Operator::MinusPlus => extra(0x2213),

        // Comparison
        Operator::Equals => symbol(0x003D),
        Operator::NotEquals => symbol(0x2260),
        Operator::LessThan => symbol(0x003C),
        Operator::GreaterThan => symbol(0x003E),
        Operator::LessThanOrEqual => symbol(0x2264),
        Operator::GreaterThanOrEqual => symbol(0x2265),

        // Algebraic
        Operator::Dot => symbol(0x22C5),
        Operator::Star => symbol(0x2217),
        Operator::Circle | Operator::Circ => extra(0x2218),
        Operator::Bullet => symbol(0x2022),
        Operator::Wedge | Operator::And => symbol(0x2227),
        Operator::Vee | Operator::Or => symbol(0x2228),
        Operator::Cap | Operator::Intersection => symbol(0x2229),
        Operator::Cup | Operator::Union => symbol(0x222A),

        // Set theory
        Operator::In => symbol(0x2208),
        Operator::NotIn => symbol(0x2209),
        Operator::Subset => symbol(0x2282),
        Operator::Superset => symbol(0x2283),
        Operator::SubsetEq => symbol(0x2286),
        Operator::SupersetEq => symbol(0x2287),
        Operator::EmptySet => symbol(0x2205),

        // Relations
        Operator::Approx => symbol(0x2248),
        Operator::Cong => symbol(0x2245),
        Operator::Equiv => symbol(0x2261),
        Operator::Propto => symbol(0x221D),
        Operator::Sim => extra(58),   // MT Extra "\sim"
        Operator::Simeq => extra(59), // MT Extra "\simeq"
        Operator::Asymp => symbol(0x224D),

        // Geometry
        Operator::Parallel => symbol(0x2225),
        Operator::Perpendicular => symbol(0x22A5),
        Operator::Angle => symbol(0x2220),

        // Calculus
        Operator::Nabla => symbol(0x2207),
        Operator::Partial => symbol(0x2202),
        Operator::Differential => CharCode::new(TYPEFACE_VARIABLE, u16::from(b'd')),

        // Special symbols
        Operator::Infinity => symbol(0x221E),
        Operator::Aleph => symbol(0x2135),
        Operator::Prime => symbol(0x2032),
        Operator::DoublePrime => symbol(0x2033),
        Operator::TriplePrime => symbol(0x2034),

        // Dots
        Operator::Ellipsis | Operator::Ldots => extra(0x2026),
        Operator::CDots => extra(0x22EF),
        Operator::VDots => extra(0x22EE),
        Operator::DDots => extra(0x22F1),

        // Arrows
        Operator::LeftArrow => symbol(0x2190),
        Operator::RightArrow => symbol(0x2192),
        Operator::UpArrow => symbol(0x2191),
        Operator::DownArrow => symbol(0x2193),
        Operator::LeftRightArrow => symbol(0x2194),
        Operator::UpDownArrow => extra(0x2195),

        // Logic
        Operator::ForAll => symbol(0x2200),
        Operator::Exists => symbol(0x2203),
        Operator::Not => symbol(216), // Symbol font "\neg"
        Operator::Implies => symbol(0x21D2),
        Operator::Iff => symbol(0x21D4),

        // Miscellaneous
        Operator::Therefore => symbol(0x2234),
        Operator::Because => extra(0x2235),
        Operator::Diamond => symbol(0x22C4),
        Operator::Box | Operator::Square => symbol(0x25A1),
    }
}

/// Map a predefined symbol to its MTEF character
pub(super) fn predefined_symbol_code(symbol: PredefinedSymbol) -> CharCode {
    /// Shorthand for a lowercase Greek character
    const fn lower(offset: u16) -> CharCode {
        CharCode::new(TYPEFACE_LCGREEK, GREEK_LOWER_FIRST as u16 + offset)
    }
    /// Shorthand for an uppercase Greek character
    const fn upper(offset: u16) -> CharCode {
        CharCode::new(TYPEFACE_UCGREEK, GREEK_UPPER_FIRST as u16 + offset)
    }

    match symbol {
        PredefinedSymbol::Alpha => lower(0),
        PredefinedSymbol::Beta => lower(1),
        PredefinedSymbol::Gamma => lower(2),
        PredefinedSymbol::Delta => lower(3),
        PredefinedSymbol::Epsilon => lower(4),
        PredefinedSymbol::Zeta => lower(5),
        PredefinedSymbol::Eta => lower(6),
        PredefinedSymbol::Theta => lower(7),
        PredefinedSymbol::Iota => lower(8),
        PredefinedSymbol::Kappa => lower(9),
        PredefinedSymbol::Lambda => lower(10),
        PredefinedSymbol::Mu => lower(11),
        PredefinedSymbol::Nu => lower(12),
        PredefinedSymbol::Xi => lower(13),
        PredefinedSymbol::Omicron => lower(14),
        PredefinedSymbol::Pi => lower(15),
        PredefinedSymbol::Rho => lower(16),
        // U+03C2 (final sigma) sits between rho and sigma, hence the gap.
        PredefinedSymbol::Sigma => lower(18),
        PredefinedSymbol::Tau => lower(19),
        PredefinedSymbol::Upsilon => lower(20),
        PredefinedSymbol::Phi => lower(21),
        PredefinedSymbol::Chi => lower(22),
        PredefinedSymbol::Psi => lower(23),
        PredefinedSymbol::Omega => lower(24),

        PredefinedSymbol::AlphaCap => upper(0),
        PredefinedSymbol::BetaCap => upper(1),
        PredefinedSymbol::GammaCap => upper(2),
        PredefinedSymbol::DeltaCap => upper(3),
        PredefinedSymbol::EpsilonCap => upper(4),
        PredefinedSymbol::ZetaCap => upper(5),
        PredefinedSymbol::EtaCap => upper(6),
        PredefinedSymbol::ThetaCap => upper(7),
        PredefinedSymbol::IotaCap => upper(8),
        PredefinedSymbol::KappaCap => upper(9),
        PredefinedSymbol::LambdaCap => upper(10),
        PredefinedSymbol::MuCap => upper(11),
        PredefinedSymbol::NuCap => upper(12),
        PredefinedSymbol::XiCap => upper(13),
        PredefinedSymbol::OmicronCap => upper(14),
        PredefinedSymbol::PiCap => upper(15),
        PredefinedSymbol::RhoCap => upper(16),
        // U+03A2 is unassigned, so the capitals skip it as well.
        PredefinedSymbol::SigmaCap => upper(18),
        PredefinedSymbol::TauCap => upper(19),
        PredefinedSymbol::UpsilonCap => upper(20),
        PredefinedSymbol::PhiCap => upper(21),
        PredefinedSymbol::ChiCap => upper(22),
        PredefinedSymbol::PsiCap => upper(23),
        PredefinedSymbol::OmegaCap => upper(24),

        PredefinedSymbol::Aleph => CharCode::new(TYPEFACE_SYMBOL, 0x2135),
        PredefinedSymbol::EulerGamma => lower(2),
        PredefinedSymbol::ExponentialE => CharCode::new(TYPEFACE_VARIABLE, u16::from(b'e')),
        PredefinedSymbol::ImaginaryI => CharCode::new(TYPEFACE_VARIABLE, u16::from(b'i')),
        PredefinedSymbol::Infinity => CharCode::new(TYPEFACE_SYMBOL, 0x221E),
    }
}

/// Map an accent to the embellishment that draws it
///
/// MTEF has no separate glyph for a caron or a grave accent; those degrade to
/// the nearest embellishment the format defines.
pub(super) fn accent_embellishment(accent: AccentType) -> u8 {
    match accent {
        AccentType::Hat => EMB_HAT,
        AccentType::Check | AccentType::Breve => EMB_SMILE,
        AccentType::Tilde => EMB_TILDE,
        AccentType::Acute => EMB_PRIME,
        AccentType::Grave => EMB_BPRIME,
        AccentType::Dot => EMB_DOT,
        AccentType::DoubleDot => EMB_DDOT,
        AccentType::TripleDot => EMB_TDOT,
        AccentType::Bar => EMB_OBAR,
        AccentType::Vec => EMB_RARROW,
    }
}

/// Base MTCode of the fixed-width spaces in the SPACE typeface
const SPACE_MTCODE_BASE: u16 = 0xEB00;

/// Map a space to its MTEF character
///
/// The SPACE typeface offers fewer widths than the AST does, so medium collapses
/// onto thick and a double quad onto a single quad.
pub(super) fn space_code(space: SpaceType) -> CharCode {
    let offset = match space {
        SpaceType::Negative => 1, // rendered as an italic correction
        SpaceType::Thin => 2,
        SpaceType::Medium | SpaceType::Thick => 4,
        SpaceType::Quad | SpaceType::QQuad => 5,
    };
    CharCode::new(TYPEFACE_SPACE, SPACE_MTCODE_BASE + offset)
}

/// Spell a predefined function the way MathType's function typeface expects
pub(super) fn function_name(function: FunctionName) -> &'static str {
    match function {
        FunctionName::Sin => "sin",
        FunctionName::Cos => "cos",
        FunctionName::Tan => "tan",
        FunctionName::Sec => "sec",
        FunctionName::Csc => "csc",
        FunctionName::Cot => "cot",
        FunctionName::ArcSin => "arcsin",
        FunctionName::ArcCos => "arccos",
        FunctionName::ArcTan => "arctan",
        FunctionName::ArcSec => "arcsec",
        FunctionName::ArcCsc => "arccsc",
        FunctionName::ArcCot => "arccot",
        FunctionName::Sinh => "sinh",
        FunctionName::Cosh => "cosh",
        FunctionName::Tanh => "tanh",
        FunctionName::Sech => "sech",
        FunctionName::Csch => "csch",
        FunctionName::Coth => "coth",
        FunctionName::Log => "log",
        FunctionName::Ln => "ln",
        FunctionName::Exp => "exp",
        FunctionName::Sqrt => "sqrt",
        FunctionName::Min => "min",
        FunctionName::Max => "max",
        FunctionName::Sup => "sup",
        FunctionName::Inf => "inf",
        FunctionName::Lim => "lim",
        FunctionName::Det => "det",
        FunctionName::Trace => "tr",
        FunctionName::Dim => "dim",
        FunctionName::Ker => "ker",
        FunctionName::Im => "Im",
        FunctionName::Re => "Re",
        FunctionName::Arg => "arg",
        FunctionName::Mod => "mod",
        FunctionName::Gcd => "gcd",
        FunctionName::Lcm => "lcm",
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::mtef::binary::charset::lookup_character;

    /// Resolve a character the way the reader would
    fn resolve(code: CharCode) -> Option<&'static str> {
        lookup_character(usize::from(code.typeface), code.mtcode, MA_MATH)
    }

    #[test]
    fn ascii_letters_use_the_variable_slot() {
        let code = char_code('x').expect("encodable");
        assert_eq!(code, CharCode::new(TYPEFACE_VARIABLE, u16::from(b'x')));
    }

    #[test]
    fn digits_use_the_number_slot() {
        let code = char_code('7').expect("encodable");
        assert_eq!(code, CharCode::new(TYPEFACE_NUMBER, u16::from(b'7')));
    }

    #[test]
    fn greek_letters_resolve_through_the_reader_tables() {
        assert_eq!(
            resolve(char_code('α').expect("encodable")),
            Some("\\alpha ")
        );
        assert_eq!(
            resolve(char_code('Ω').expect("encodable")),
            Some("\\Omega ")
        );
    }

    #[test]
    fn non_bmp_characters_are_rejected() {
        assert_eq!(
            char_code('𝕏'),
            Err(MtefWriteError::UnsupportedCharacter('𝕏'))
        );
    }

    #[test]
    fn every_operator_resolves_or_falls_back_to_its_codepoint() {
        // Operators whose glyph the reader has no table entry for still round-trip
        // because the reader falls back to the raw codepoint.
        let cases = [
            (Operator::Plus, "+"),
            (Operator::Minus, "-"),
            (Operator::Equals, "="),
            (Operator::Times, "\\times "),
            (Operator::Divide, "\\div "),
            (Operator::PlusMinus, "\\pm "),
            (Operator::LessThanOrEqual, "\\leq "),
            (Operator::GreaterThanOrEqual, "\\geq "),
            (Operator::NotEquals, "\\neq "),
            (Operator::Approx, "\\approx "),
            (Operator::Equiv, "\\equiv "),
            (Operator::In, "\\in "),
            (Operator::Infinity, "\\infty "),
            (Operator::Partial, "\\partial "),
            (Operator::Nabla, "\\nabla "),
            (Operator::RightArrow, "\\rightarrow "),
            (Operator::Ellipsis, "\\ldots "),
            (Operator::CDots, "\\cdots "),
            (Operator::Not, "\\neg "),
            (Operator::Therefore, "\\therefore "),
        ];
        for (operator, expected) in cases {
            assert_eq!(
                resolve(operator_code(operator)),
                Some(expected),
                "{operator:?}"
            );
        }
    }

    #[test]
    fn predefined_greek_symbols_resolve() {
        assert_eq!(
            resolve(predefined_symbol_code(PredefinedSymbol::Sigma)),
            Some("\\sigma ")
        );
        assert_eq!(
            resolve(predefined_symbol_code(PredefinedSymbol::SigmaCap)),
            Some("\\Sigma ")
        );
        assert_eq!(
            resolve(predefined_symbol_code(PredefinedSymbol::Omega)),
            Some("\\omega ")
        );
        assert_eq!(
            resolve(predefined_symbol_code(PredefinedSymbol::Pi)),
            Some("\\pi ")
        );
    }

    #[test]
    fn spaces_resolve_to_latex_spacing_commands() {
        assert_eq!(resolve(space_code(SpaceType::Thin)), Some("\\,"));
        assert_eq!(resolve(space_code(SpaceType::Thick)), Some("\\;"));
        assert_eq!(resolve(space_code(SpaceType::Quad)), Some("\\quad "));
    }
}
