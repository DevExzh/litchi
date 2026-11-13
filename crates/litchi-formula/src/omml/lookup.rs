use super::elements::ElementType;
use crate::ast::{AccentType, LargeOperator, StyleType};
use phf::{phf_map, phf_set};

/// Fast element name to type lookup using PHF
pub static ELEMENT_TYPES: phf::Map<&'static str, ElementType> = phf_map! {
    // Core math elements
    "oMath" => ElementType::Math,
    "m:oMath" => ElementType::Math,

    // Text and runs
    "r" => ElementType::Run,
    "m:r" => ElementType::Run,
    "t" => ElementType::Text,
    "m:t" => ElementType::Text,

    // Functions
    "func" => ElementType::Function,
    "m:func" => ElementType::Function,
    "fName" => ElementType::FunctionName,
    "m:fName" => ElementType::FunctionName,

    // Fractions
    "f" => ElementType::Fraction,
    "m:f" => ElementType::Fraction,
    "num" => ElementType::Numerator,
    "m:num" => ElementType::Numerator,
    "den" => ElementType::Denominator,
    "m:den" => ElementType::Denominator,

    // Radicals
    "rad" => ElementType::Radical,
    "m:rad" => ElementType::Radical,
    "deg" => ElementType::Degree,
    "m:deg" => ElementType::Degree,

    // Scripts
    "sSup" => ElementType::Superscript,
    "m:sSup" => ElementType::Superscript,
    "sSub" => ElementType::Subscript,
    "m:sSub" => ElementType::Subscript,
    "sSubSup" => ElementType::SubSup,
    "m:sSubSup" => ElementType::SubSup,
    "sPre" => ElementType::PreScript,
    "m:sPre" => ElementType::PreScript,
    "sPost" => ElementType::PostScript,
    "m:sPost" => ElementType::PostScript,
    "sup" => ElementType::SuperscriptElement,
    "m:sup" => ElementType::SuperscriptElement,
    "sub" => ElementType::SubscriptElement,
    "m:sub" => ElementType::SubscriptElement,

    // Delimiters
    "d" => ElementType::Delimiter,
    "m:d" => ElementType::Delimiter,

    // N-ary operators
    "nary" => ElementType::Nary,
    "m:nary" => ElementType::Nary,
    "lim" => ElementType::Limit,
    "m:lim" => ElementType::Limit,
    "limLow" => ElementType::LimLow,
    "m:limLow" => ElementType::LimLow,
    "limUpp" => ElementType::LimUpp,
    "m:limUpp" => ElementType::LimUpp,

    // Matrices
    "m" => ElementType::Matrix,
    "m:m" => ElementType::Matrix,
    "mr" => ElementType::MatrixRow,
    "m:mr" => ElementType::MatrixRow,

    // Equation arrays
    "eqArr" => ElementType::EqArr,
    "m:eqArr" => ElementType::EqArr,

    // Accents and decorations
    "acc" => ElementType::Accent,
    "m:acc" => ElementType::Accent,
    "bar" => ElementType::Bar,
    "m:bar" => ElementType::Bar,
    "box" => ElementType::Box,
    "m:box" => ElementType::Box,
    "phant" => ElementType::Phantom,
    "m:phant" => ElementType::Phantom,
    "groupChr" => ElementType::GroupChar,
    "m:groupChr" => ElementType::GroupChar,
    "borderBox" => ElementType::BorderBox,
    "m:borderBox" => ElementType::BorderBox,

    // Properties and styles
    "rPr" => ElementType::Properties,
    "m:rPr" => ElementType::Properties,
    "lit" => ElementType::Lit,
    "m:lit" => ElementType::Lit,
    "scr" => ElementType::Scr,
    "m:scr" => ElementType::Scr,
    "sty" => ElementType::Sty,
    "m:sty" => ElementType::Sty,
    "nor" => ElementType::Nor,
    "m:nor" => ElementType::Nor,
    "fPr" => ElementType::Properties,
    "m:fPr" => ElementType::Properties,
    "dPr" => ElementType::Properties,
    "m:dPr" => ElementType::Properties,
    "naryPr" => ElementType::Properties,
    "m:naryPr" => ElementType::Properties,
    "accPr" => ElementType::AccentProperties,
    "m:accPr" => ElementType::AccentProperties,
    "radPr" => ElementType::Properties,
    "m:radPr" => ElementType::Properties,
    "sSupPr" => ElementType::Properties,
    "m:sSupPr" => ElementType::Properties,
    "sSubPr" => ElementType::Properties,
    "m:sSubPr" => ElementType::Properties,
    "funcPr" => ElementType::Properties,
    "m:funcPr" => ElementType::Properties,
    "groupChrPr" => ElementType::Properties,
    "m:groupChrPr" => ElementType::Properties,
    "eqArrPr" => ElementType::EqArrPr,
    "m:eqArrPr" => ElementType::EqArrPr,
    "boxPr" => ElementType::Properties,
    "m:boxPr" => ElementType::Properties,
    "phantPr" => ElementType::Properties,
    "m:phantPr" => ElementType::Properties,
    "borderBoxPr" => ElementType::Properties,
    "m:borderBoxPr" => ElementType::Properties,
    "mPr" => ElementType::Properties,
    "m:mPr" => ElementType::Properties,

    // Characters and content
    "chr" => ElementType::Character,
    "m:chr" => ElementType::Character,
    "pos" => ElementType::Position,
    "m:pos" => ElementType::Position,
    "vertJc" => ElementType::VerticalAlignment,
    "m:vertJc" => ElementType::VerticalAlignment,
    "begChr" => ElementType::Character,
    "m:begChr" => ElementType::Character,
    "endChr" => ElementType::Character,
    "m:endChr" => ElementType::Character,
    "sepChr" => ElementType::Character,
    "m:sepChr" => ElementType::Character,
    "e" => ElementType::Base,
    "m:e" => ElementType::Base,

    // Single-value property elements (carry an m:val attribute or text content)
    "type" => ElementType::ValProperty,
    "m:type" => ElementType::ValProperty,
    "subHide" => ElementType::ValProperty,
    "m:subHide" => ElementType::ValProperty,
    "supHide" => ElementType::ValProperty,
    "m:supHide" => ElementType::ValProperty,
    "baseJc" => ElementType::ValProperty,
    "m:baseJc" => ElementType::ValProperty,
    "maxDist" => ElementType::ValProperty,
    "m:maxDist" => ElementType::ValProperty,
    "objDist" => ElementType::ValProperty,
    "m:objDist" => ElementType::ValProperty,
    "rSp" => ElementType::ValProperty,
    "m:rSp" => ElementType::ValProperty,
    "rSpRule" => ElementType::ValProperty,
    "m:rSpRule" => ElementType::ValProperty,

    // Spacing
};

/// Fast accent character to accent type lookup
///
/// Includes both standalone accent characters and Unicode combining characters
/// commonly used in OMML (Office Math Markup Language).
///
/// Reference: LibreOffice starmath/inc/types.hxx and plurimath
pub static ACCENTS: phf::Map<&'static str, AccentType> = phf_map! {
    // Standalone accent characters
    "¯" => AccentType::Bar,      // U+00AF Macron
    "¨" => AccentType::DoubleDot, // U+00A8 Diaeresis
    "˙" => AccentType::Dot,      // U+02D9 Dot above
    "`" => AccentType::Grave,    // U+0060 Grave accent
    "´" => AccentType::Acute,    // U+00B4 Acute accent
    "˜" => AccentType::Tilde,    // U+007E Tilde
    "^" => AccentType::Hat,      // U+005E Circumflex
    "ˇ" => AccentType::Check,    // U+02C7 Caron
    "˘" => AccentType::Breve,    // U+02D8 Breve
    "→" => AccentType::Vec,      // U+2192 Right arrow
    "⃗" => AccentType::Vec,      // U+20D7 Combining right arrow above
    "⃛" => AccentType::TripleDot, // U+20DB Combining three dots above

    // Combining characters (Unicode U+0300-U+036F range)
    // These are commonly used in OMML documents
    "̀" => AccentType::Grave,    // U+0300 Combining grave accent
    "́" => AccentType::Acute,    // U+0301 Combining acute accent
    "̂" => AccentType::Hat,      // U+0302 Combining circumflex
    "̃" => AccentType::Tilde,    // U+0303 Combining tilde
    "̄" => AccentType::Bar,      // U+0304 Combining macron (bar)
    "̅" => AccentType::Bar,      // U+0305 Combining overline - COMMON IN OMML!
    "̆" => AccentType::Breve,    // U+0306 Combining breve
    "̇" => AccentType::Dot,      // U+0307 Combining dot above
    "̈" => AccentType::DoubleDot, // U+0308 Combining diaeresis
    "̊" => AccentType::Dot,      // U+030A Combining ring above (treated as dot)
    "̌" => AccentType::Check,    // U+030C Combining caron
};

/// Fast large operator character to operator lookup
pub static LARGE_OPERATORS: phf::Map<&'static str, LargeOperator> = phf_map! {
    "∑" => LargeOperator::Sum,
    "∏" => LargeOperator::Product,
    "∐" => LargeOperator::Coproduct,
    "∫" => LargeOperator::Integral,
    "∬" => LargeOperator::DoubleIntegral,
    "∭" => LargeOperator::TripleIntegral,
    "∮" => LargeOperator::ContourIntegral,
    "∯" => LargeOperator::SurfaceIntegral,
    "∰" => LargeOperator::VolumeIntegral,
    "⋃" => LargeOperator::Union,
    "⋂" => LargeOperator::Intersection,
    "⨄" => LargeOperator::BigUnion,
    "⨅" => LargeOperator::BigIntersection,
    "lim" => LargeOperator::Limit,
    "max" => LargeOperator::Max,
    "min" => LargeOperator::Min,
    "sup" => LargeOperator::Supremum,
    "inf" => LargeOperator::Infimum,
    "argmax" => LargeOperator::ArgMax,
    "argmin" => LargeOperator::ArgMin,
};

/// Fast attribute value parsing lookup tables
pub static BOOLEAN_VALUES: phf::Set<&'static str> = phf_set! {
    "1", "true", "on", "yes"
};

pub static STYLE_VALUES: phf::Map<&'static str, StyleType> = phf_map! {
    "p" => StyleType::Normal,
    "roman" => StyleType::Normal,
    "b" => StyleType::Bold,
    "bold" => StyleType::Bold,
    "i" => StyleType::Italic,
    "italic" => StyleType::Italic,
    "bi" => StyleType::BoldItalic,
    "bold-italic" => StyleType::BoldItalic,
    "ss" => StyleType::SansSerif,
    "sans-serif" => StyleType::SansSerif,
    "ssb" => StyleType::SansSerifBold,
    "sans-serif-bold" => StyleType::SansSerifBold,
    "ssi" => StyleType::SansSerifItalic,
    "sans-serif-italic" => StyleType::SansSerifItalic,
    "ssbi" => StyleType::SansSerifBoldItalic,
    "sans-serif-bold-italic" => StyleType::SansSerifBoldItalic,
    "m" => StyleType::Monospace,
    "monospace" => StyleType::Monospace,
    "sc" => StyleType::Script,
    "script" => StyleType::Script,
    "bsc" => StyleType::BoldScript,
    "bold-script" => StyleType::BoldScript,
    "fr" => StyleType::Fraktur,
    "fraktur" => StyleType::Fraktur,
    "bfr" => StyleType::BoldFraktur,
    "bold-fraktur" => StyleType::BoldFraktur,
    "ds" => StyleType::DoubleStruck,
    "double-struck" => StyleType::DoubleStruck,
};

/// Fast lookup functions using PHF maps
/// Get element type from name with fallback to Unknown
pub fn get_element_type(name: &str) -> ElementType {
    ELEMENT_TYPES
        .get(name)
        .copied()
        .unwrap_or(ElementType::Unknown)
}

/// Get accent type from character
pub fn get_accent_type(chr: &str) -> Option<AccentType> {
    ACCENTS.get(chr).copied()
}

/// Get large operator from character
pub fn get_large_operator(chr: &str) -> Option<LargeOperator> {
    LARGE_OPERATORS.get(chr).copied()
}

/// Parse boolean attribute value
pub fn parse_bool_value(val: &str) -> Option<bool> {
    if BOOLEAN_VALUES.contains(val) {
        Some(true)
    } else if val == "0" || val == "false" || val == "off" || val == "no" {
        Some(false)
    } else {
        None
    }
}

/// Parse style value
pub fn parse_style_value(val: &str) -> Option<StyleType> {
    STYLE_VALUES.get(val).copied()
}
