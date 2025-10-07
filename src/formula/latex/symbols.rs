// Symbol conversion to LaTeX
//
// This module handles conversion of mathematical symbols to LaTeX format,
// including Greek letters, special symbols, and Unicode character mapping.
// Uses efficient lookup tables and SIMD operations for performance.

use crate::formula::ast::Symbol;
use super::LatexError;

// Static lookup table for common Greek letters and symbols
static GREEK_SYMBOLS: phf::Map<&'static str, &'static str> = phf::phf_map! {
    // Lowercase Greek
    "alpha" => "\\alpha",
    "beta" => "\\beta",
    "gamma" => "\\gamma",
    "delta" => "\\delta",
    "epsilon" => "\\epsilon",
    "zeta" => "\\zeta",
    "eta" => "\\eta",
    "theta" => "\\theta",
    "iota" => "\\iota",
    "kappa" => "\\kappa",
    "lambda" => "\\lambda",
    "mu" => "\\mu",
    "nu" => "\\nu",
    "xi" => "\\xi",
    "omicron" => "o",
    "pi" => "\\pi",
    "rho" => "\\rho",
    "sigma" => "\\sigma",
    "tau" => "\\tau",
    "upsilon" => "\\upsilon",
    "phi" => "\\phi",
    "chi" => "\\chi",
    "psi" => "\\psi",
    "omega" => "\\omega",

    // Uppercase Greek
    "Alpha" => "A",
    "Beta" => "B",
    "Gamma" => "\\Gamma",
    "Delta" => "\\Delta",
    "Epsilon" => "E",
    "Zeta" => "Z",
    "Eta" => "H",
    "Theta" => "\\Theta",
    "Iota" => "I",
    "Kappa" => "K",
    "Lambda" => "\\Lambda",
    "Mu" => "M",
    "Nu" => "N",
    "Xi" => "\\Xi",
    "Omicron" => "O",
    "Pi" => "\\Pi",
    "Rho" => "P",
    "Sigma" => "\\Sigma",
    "Tau" => "T",
    "Upsilon" => "\\Upsilon",
    "Phi" => "\\Phi",
    "Chi" => "X",
    "Psi" => "\\Psi",
    "Omega" => "\\Omega",

    // Special symbols
    "aleph" => "\\aleph",
    "hbar" => "\\hbar",
    "ell" => "\\ell",
    "wp" => "\\wp",
    "Re" => "\\Re",
    "Im" => "\\Im",
    "prime" => "'",
    "doubleprime" => "''",
    "partial" => "\\partial",
    "infty" => "\\infty",
    "emptyset" => "\\emptyset",
    "nabla" => "\\nabla",
    "surd" => "\\surd",
    "top" => "\\top",
    "bot" => "\\bot",
    "angle" => "\\angle",
    "triangle" => "\\triangle",
    "backslash" => "\\backslash",
    "forall" => "\\forall",
    "exists" => "\\exists",
    "nexists" => "\\nexists",
    "therefore" => "\\therefore",
    "because" => "\\because",
    "mapsto" => "\\mapsto",
    "to" => "\\to",
    "gets" => "\\gets",
    "leftrightarrow" => "\\leftrightarrow",
    "uparrow" => "\\uparrow",
    "downarrow" => "\\downarrow",
    "updownarrow" => "\\updownarrow",
    "Uparrow" => "\\Uparrow",
    "Downarrow" => "\\Downarrow",
    "Updownarrow" => "\\Updownarrow",
    "nearrow" => "\\nearrow",
    "searrow" => "\\searrow",
    "swarrow" => "\\swarrow",
    "nwarrow" => "\\nwarrow",
    "leftharpoonup" => "\\leftharpoonup",
    "leftharpoondown" => "\\leftharpoondown",
    "rightharpoonup" => "\\rightharpoonup",
    "rightharpoondown" => "\\rightharpoondown",
    "hookleftarrow" => "\\hookleftarrow",
    "hookrightarrow" => "\\hookrightarrow",
    "longleftarrow" => "\\longleftarrow",
    "longrightarrow" => "\\longrightarrow",
    "Longleftarrow" => "\\Longleftarrow",
    "Longrightarrow" => "\\Longrightarrow",
    "longleftrightarrow" => "\\longleftrightarrow",
    "Longleftrightarrow" => "\\Longleftrightarrow",
};

/// Unicode to LaTeX mapping for common mathematical symbols
static UNICODE_TO_LATEX: phf::Map<char, &'static str> = phf::phf_map! {
    // Greek letters
    'α' => "\\alpha",
    'β' => "\\beta",
    'γ' => "\\gamma",
    'δ' => "\\delta",
    'ε' => "\\epsilon",
    'ζ' => "\\zeta",
    'η' => "\\eta",
    'θ' => "\\theta",
    'ι' => "\\iota",
    'κ' => "\\kappa",
    'λ' => "\\lambda",
    'μ' => "\\mu",
    'ν' => "\\nu",
    'ξ' => "\\xi",
    'π' => "\\pi",
    'ρ' => "\\rho",
    'σ' => "\\sigma",
    'τ' => "\\tau",
    'υ' => "\\upsilon",
    'φ' => "\\phi",
    'χ' => "\\chi",
    'ψ' => "\\psi",
    'ω' => "\\omega",

    // Uppercase Greek
    'Γ' => "\\Gamma",
    'Δ' => "\\Delta",
    'Θ' => "\\Theta",
    'Λ' => "\\Lambda",
    'Ξ' => "\\Xi",
    'Π' => "\\Pi",
    'Σ' => "\\Sigma",
    'Υ' => "\\Upsilon",
    'Φ' => "\\Phi",
    'Ψ' => "\\Psi",
    'Ω' => "\\Omega",

    // Operators and symbols
    '∑' => "\\sum",
    '∏' => "\\prod",
    '∫' => "\\int",
    '∮' => "\\oint",
    '√' => "\\sqrt",
    '∂' => "\\partial",
    '∇' => "\\nabla",
    '∞' => "\\infty",
    '∅' => "\\emptyset",
    '∀' => "\\forall",
    '∃' => "\\exists",
    '∄' => "\\nexists",
    '∴' => "\\therefore",
    '∵' => "\\because",
    '⊂' => "\\subset",
    '⊃' => "\\supset",
    '⊆' => "\\subseteq",
    '⊇' => "\\supseteq",
    '∈' => "\\in",
    '∉' => "\\notin",
    '∩' => "\\cap",
    '∪' => "\\cup",
    '≠' => "\\neq",
    '≤' => "\\leq",
    '≥' => "\\geq",
    '≈' => "\\approx",
    '∼' => "\\sim",
    '∝' => "\\propto",
    '≡' => "\\equiv",
    '≪' => "\\ll",
    '≫' => "\\gg",
    '→' => "\\to",
    '←' => "\\gets",
    '↔' => "\\leftrightarrow",
    '↑' => "\\uparrow",
    '↓' => "\\downarrow",
    '⇑' => "\\Uparrow",
    '⇓' => "\\Downarrow",
    '⇒' => "\\Rightarrow",
    '⇐' => "\\Leftarrow",
    '⇔' => "\\Leftrightarrow",
    '±' => "\\pm",
    '∓' => "\\mp",
    '×' => "\\times",
    '÷' => "\\div",
    '⋅' => "\\cdot",
    '∗' => "\\ast",
    '∘' => "\\circ",
    '∧' => "\\wedge",
    '∨' => "\\vee",
    '⊕' => "\\oplus",
    '⊗' => "\\otimes",
    '⊙' => "\\odot",
    '△' => "\\triangle",
    '□' => "\\square",
    '◇' => "\\diamond",
    '†' => "\\dagger",
    '‡' => "\\ddagger",
    '‾' => "\\bar",
    'ˆ' => "\\hat",
    '˜' => "\\tilde",
    'ˇ' => "\\check",
    '´' => "\\acute",
    '˙' => "\\dot",
    '¨' => "\\ddot",
    '˘' => "\\breve",
    '°' => "\\degree",
    '′' => "'",
    '″' => "''",
    '‴' => "'''",
    'ℵ' => "\\aleph",
    'ℶ' => "\\beth",
    'ℷ' => "\\gimel",
    'ℸ' => "\\daleth",
    'ℏ' => "\\hbar",
    'ℓ' => "\\ell",
    '℘' => "\\wp",
    'ℜ' => "\\Re",
    'ℑ' => "\\Im",
    'ℕ' => "\\mathbb{N}",
    'ℤ' => "\\mathbb{Z}",
    'ℚ' => "\\mathbb{Q}",
    'ℝ' => "\\mathbb{R}",
    'ℂ' => "\\mathbb{C}",
    'ℍ' => "\\mathbb{H}",
    'ℙ' => "\\mathbb{P}",
};

/// Convert a symbol to LaTeX format
///
/// Handles both named symbols (Greek letters, special symbols) and Unicode characters.
/// Uses efficient lookup tables for performance.
pub fn convert_symbol(buffer: &mut String, symbol: &Symbol) -> Result<(), LatexError> {
    // If Unicode character is provided, try Unicode lookup first
    if let Some(unicode) = symbol.unicode {
        if let Some(latex) = UNICODE_TO_LATEX.get(&unicode) {
            buffer.push_str(latex);
            return Ok(());
        }
        // Fall back to using the Unicode character directly if not in our mapping
        buffer.push(unicode);
        return Ok(());
    }

    // Try named symbol lookup
    if let Some(latex) = GREEK_SYMBOLS.get(symbol.name.as_ref()) {
        buffer.push_str(latex);
        return Ok(());
    }

    // If no mapping found, use the name directly
    // For unknown symbols, wrap in \text{} to ensure proper rendering
    buffer.push_str("\\text{");
    buffer.push_str(&symbol.name);
    buffer.push('}');

    Ok(())
}


#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::ast::Symbol;
    use std::borrow::Cow;

    #[test]
    fn test_convert_greek_symbol() {
        let mut buffer = String::new();
        let symbol = Symbol {
            name: Cow::Borrowed("alpha"),
            unicode: None,
            variant: None,
        };

        convert_symbol(&mut buffer, &symbol).unwrap();
        assert_eq!(buffer, "\\alpha");
    }

    #[test]
    fn test_convert_unicode_greek() {
        let mut buffer = String::new();
        let symbol = Symbol {
            name: Cow::Borrowed("beta"),
            unicode: Some('β'),
            variant: None,
        };

        convert_symbol(&mut buffer, &symbol).unwrap();
        assert_eq!(buffer, "\\beta");
    }

    #[test]
    fn test_convert_unknown_symbol() {
        let mut buffer = String::new();
        let symbol = Symbol {
            name: Cow::Borrowed("unknown_symbol"),
            unicode: None,
            variant: None,
        };

        convert_symbol(&mut buffer, &symbol).unwrap();
        assert_eq!(buffer, "\\text{unknown_symbol}");
    }

    #[test]
    fn test_unicode_fallback() {
        let mut buffer = String::new();
        let symbol = Symbol {
            name: Cow::Borrowed("custom"),
            unicode: Some('©'), // Copyright symbol not in our mapping
            variant: None,
        };

        convert_symbol(&mut buffer, &symbol).unwrap();
        assert_eq!(buffer, "©");
    }
}
