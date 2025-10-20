//! Unicode superscript and subscript character conversion.
//!
//! This module provides zero-cost compile-time lookup tables for converting
//! regular characters to their Unicode superscript and subscript equivalents.
//! Uses `phf` for efficient perfect hash function lookups.
//!
//! # Unicode Character Coverage
//!
//! ## Superscripts
//! - Digits: 0-9 → ⁰¹²³⁴⁵⁶⁷⁸⁹
//! - Latin letters: ⁱⁿᵃᵇᶜᵈᵉᶠᵍʰʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻ
//! - Greek letters: ᵝᵞᵟᵠᵡ
//! - Symbols: ⁺⁻⁼⁽⁾
//!
//! ## Subscripts
//! - Digits: 0-9 → ₀₁₂₃₄₅₆₇₈₉
//! - Latin letters: ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ
//! - Symbols: ₊₋₌₍₎
//!
//! # Examples
//!
//! ```rust
//! use litchi::markdown::unicode::{to_superscript, to_subscript};
//!
//! // Convert single character
//! assert_eq!(to_superscript('2'), Some('²'));
//! assert_eq!(to_subscript('0'), Some('₀'));
//!
//! // Convert string (note: 'x' also has superscript)
//! let superscript = "x2".chars().map(|c| to_superscript(c).unwrap_or(c)).collect::<String>();
//! assert_eq!(superscript, "ˣ²");
//! ```
use phf::phf_map;

/// Compile-time lookup table for superscript characters.
///
/// Maps regular characters to their Unicode superscript equivalents.
/// Uses perfect hash function for O(1) lookup with zero runtime cost.
static SUPERSCRIPT_MAP: phf::Map<char, char> = phf_map! {
    // Digits
    '0' => '⁰',
    '1' => '¹',
    '2' => '²',
    '3' => '³',
    '4' => '⁴',
    '5' => '⁵',
    '6' => '⁶',
    '7' => '⁷',
    '8' => '⁸',
    '9' => '⁹',
    
    // Latin lowercase letters
    'a' => 'ᵃ',
    'b' => 'ᵇ',
    'c' => 'ᶜ',
    'd' => 'ᵈ',
    'e' => 'ᵉ',
    'f' => 'ᶠ',
    'g' => 'ᵍ',
    'h' => 'ʰ',
    'i' => 'ⁱ',
    'j' => 'ʲ',
    'k' => 'ᵏ',
    'l' => 'ˡ',
    'm' => 'ᵐ',
    'n' => 'ⁿ',
    'o' => 'ᵒ',
    'p' => 'ᵖ',
    'r' => 'ʳ',
    's' => 'ˢ',
    't' => 'ᵗ',
    'u' => 'ᵘ',
    'v' => 'ᵛ',
    'w' => 'ʷ',
    'x' => 'ˣ',
    'y' => 'ʸ',
    'z' => 'ᶻ',
    
    // Latin uppercase letters (limited support in Unicode)
    'A' => 'ᴬ',
    'B' => 'ᴮ',
    'D' => 'ᴰ',
    'E' => 'ᴱ',
    'G' => 'ᴳ',
    'H' => 'ᴴ',
    'I' => 'ᴵ',
    'J' => 'ᴶ',
    'K' => 'ᴷ',
    'L' => 'ᴸ',
    'M' => 'ᴹ',
    'N' => 'ᴺ',
    'O' => 'ᴼ',
    'P' => 'ᴾ',
    'R' => 'ᴿ',
    'T' => 'ᵀ',
    'U' => 'ᵁ',
    'V' => 'ᵛ',
    'W' => 'ᵂ',
    
    // Greek letters
    'β' => 'ᵝ',
    'γ' => 'ᵞ',
    'δ' => 'ᵟ',
    'φ' => 'ᵠ',
    'χ' => 'ᵡ',
    
    // Symbols
    '+' => '⁺',
    '-' => '⁻',
    '=' => '⁼',
    '(' => '⁽',
    ')' => '⁾',
};

/// Compile-time lookup table for subscript characters.
///
/// Maps regular characters to their Unicode subscript equivalents.
/// Uses perfect hash function for O(1) lookup with zero runtime cost.
static SUBSCRIPT_MAP: phf::Map<char, char> = phf_map! {
    // Digits
    '0' => '₀',
    '1' => '₁',
    '2' => '₂',
    '3' => '₃',
    '4' => '₄',
    '5' => '₅',
    '6' => '₆',
    '7' => '₇',
    '8' => '₈',
    '9' => '₉',
    
    // Latin lowercase letters (limited support)
    'a' => 'ₐ',
    'e' => 'ₑ',
    'h' => 'ₕ',
    'i' => 'ᵢ',
    'j' => 'ⱼ',
    'k' => 'ₖ',
    'l' => 'ₗ',
    'm' => 'ₘ',
    'n' => 'ₙ',
    'o' => 'ₒ',
    'p' => 'ₚ',
    'r' => 'ᵣ',
    's' => 'ₛ',
    't' => 'ₜ',
    'u' => 'ᵤ',
    'v' => 'ᵥ',
    'x' => 'ₓ',
    
    // Greek letters
    'β' => 'ᵦ',
    'γ' => 'ᵧ',
    'ρ' => 'ᵨ',
    'φ' => 'ᵩ',
    'χ' => 'ᵪ',
    
    // Symbols
    '+' => '₊',
    '-' => '₋',
    '=' => '₌',
    '(' => '₍',
    ')' => '₎',
};

/// Convert a character to its Unicode superscript equivalent.
///
/// Returns `Some(char)` if a superscript equivalent exists, `None` otherwise.
///
/// # Examples
///
/// ```rust
/// use litchi::markdown::unicode::to_superscript;
///
/// assert_eq!(to_superscript('2'), Some('²'));
/// assert_eq!(to_superscript('n'), Some('ⁿ'));
/// assert_eq!(to_superscript('+'), Some('⁺'));
/// assert_eq!(to_superscript('q'), None); // No Unicode superscript for 'q'
/// ```
///
/// # Performance
///
/// This function uses a compile-time perfect hash function for O(1) lookup
/// with zero runtime cost. The lookup table is embedded directly in the binary.
#[inline]
pub fn to_superscript(c: char) -> Option<char> {
    SUPERSCRIPT_MAP.get(&c).copied()
}

/// Convert a character to its Unicode subscript equivalent.
///
/// Returns `Some(char)` if a subscript equivalent exists, `None` otherwise.
///
/// # Examples
///
/// ```rust
/// use litchi::markdown::unicode::to_subscript;
///
/// assert_eq!(to_subscript('0'), Some('₀'));
/// assert_eq!(to_subscript('i'), Some('ᵢ'));
/// assert_eq!(to_subscript('+'), Some('₊'));
/// assert_eq!(to_subscript('b'), None); // No Unicode subscript for 'b'
/// ```
///
/// # Performance
///
/// This function uses a compile-time perfect hash function for O(1) lookup
/// with zero runtime cost. The lookup table is embedded directly in the binary.
#[inline]
pub fn to_subscript(c: char) -> Option<char> {
    SUBSCRIPT_MAP.get(&c).copied()
}

/// Convert a string to superscript, falling back to original characters for unsupported ones.
///
/// This function attempts to convert each character in the input string to its
/// superscript equivalent. Characters without a superscript mapping remain unchanged.
///
/// # Examples
///
/// ```rust
/// use litchi::markdown::unicode::convert_to_superscript;
///
/// // 'x' has superscript, so it converts too
/// assert_eq!(convert_to_superscript("x2"), "ˣ²");
/// assert_eq!(convert_to_superscript("n+1"), "ⁿ⁺¹");
/// assert_eq!(convert_to_superscript("2nd"), "²ⁿᵈ");
/// // Some uppercase letters have superscripts
/// assert_eq!(convert_to_superscript("H2O"), "ᴴ²ᴼ");
/// // Characters without superscript remain unchanged (e.g., 'C')
/// assert_eq!(convert_to_superscript("CO2"), "Cᴼ²");
/// ```
///
/// # Performance
///
/// This function pre-allocates the output string with the exact capacity needed,
/// minimizing allocations. Character conversion uses zero-cost lookups.
#[inline]
pub fn convert_to_superscript(text: &str) -> String {
    // Pre-allocate with same capacity as input (superscript chars are same byte size or larger)
    let mut result = String::with_capacity(text.len() * 2);
    
    for c in text.chars() {
        result.push(to_superscript(c).unwrap_or(c));
    }
    
    result
}

/// Convert a string to subscript, falling back to original characters for unsupported ones.
///
/// This function attempts to convert each character in the input string to its
/// subscript equivalent. Characters without a subscript mapping remain unchanged.
///
/// # Examples
///
/// ```rust
/// use litchi::markdown::unicode::convert_to_subscript;
///
/// // Uppercase 'H' and 'O' have no subscript (only lowercase letters have subscripts)
/// assert_eq!(convert_to_subscript("H2O"), "H₂O");
/// // Lowercase letters with subscripts convert
/// assert_eq!(convert_to_subscript("h2o"), "ₕ₂ₒ");
/// // 'x' has subscript, so it converts too
/// assert_eq!(convert_to_subscript("x0"), "ₓ₀");
/// assert_eq!(convert_to_subscript("a+b"), "ₐ₊b"); // 'b' has no subscript
/// ```
///
/// # Performance
///
/// This function pre-allocates the output string with the exact capacity needed,
/// minimizing allocations. Character conversion uses zero-cost lookups.
#[inline]
pub fn convert_to_subscript(text: &str) -> String {
    // Pre-allocate with same capacity as input (subscript chars are same byte size or larger)
    let mut result = String::with_capacity(text.len() * 2);
    
    for c in text.chars() {
        result.push(to_subscript(c).unwrap_or(c));
    }
    
    result
}

/// Check if all characters in a string can be converted to superscript.
///
/// Returns `true` if all characters have Unicode superscript equivalents.
///
/// # Examples
///
/// ```rust
/// use litchi::markdown::unicode::can_convert_to_superscript;
///
/// assert!(can_convert_to_superscript("123"));
/// assert!(can_convert_to_superscript("n+1"));
/// assert!(!can_convert_to_superscript("query")); // 'q' has no superscript
/// ```
#[inline]
pub fn can_convert_to_superscript(text: &str) -> bool {
    text.chars().all(|c| to_superscript(c).is_some())
}

/// Check if all characters in a string can be converted to subscript.
///
/// Returns `true` if all characters have Unicode subscript equivalents.
///
/// # Examples
///
/// ```rust
/// use litchi::markdown::unicode::can_convert_to_subscript;
///
/// assert!(can_convert_to_subscript("123"));
/// assert!(can_convert_to_subscript("i+1"));
/// assert!(!can_convert_to_subscript("abc")); // 'b' and 'c' have no subscript
/// ```
#[inline]
pub fn can_convert_to_subscript(text: &str) -> bool {
    text.chars().all(|c| to_subscript(c).is_some())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_superscript_digits() {
        assert_eq!(to_superscript('0'), Some('⁰'));
        assert_eq!(to_superscript('1'), Some('¹'));
        assert_eq!(to_superscript('2'), Some('²'));
        assert_eq!(to_superscript('9'), Some('⁹'));
    }

    #[test]
    fn test_subscript_digits() {
        assert_eq!(to_subscript('0'), Some('₀'));
        assert_eq!(to_subscript('1'), Some('₁'));
        assert_eq!(to_subscript('2'), Some('₂'));
        assert_eq!(to_subscript('9'), Some('₉'));
    }

    #[test]
    fn test_superscript_letters() {
        assert_eq!(to_superscript('n'), Some('ⁿ'));
        assert_eq!(to_superscript('i'), Some('ⁱ'));
        assert_eq!(to_superscript('x'), Some('ˣ'));
        assert_eq!(to_superscript('q'), None); // No superscript for 'q'
    }

    #[test]
    fn test_subscript_letters() {
        assert_eq!(to_subscript('i'), Some('ᵢ'));
        assert_eq!(to_subscript('n'), Some('ₙ'));
        assert_eq!(to_subscript('b'), None); // No subscript for 'b'
    }

    #[test]
    fn test_convert_to_superscript() {
        // 'x' has a superscript equivalent 'ˣ', so it gets converted
        assert_eq!(convert_to_superscript("x2"), "ˣ²");
        assert_eq!(convert_to_superscript("n+1"), "ⁿ⁺¹");
        assert_eq!(convert_to_superscript("2nd"), "²ⁿᵈ");
        // Uppercase H and O have superscript forms, so they convert
        assert_eq!(convert_to_superscript("H2O"), "ᴴ²ᴼ");
        assert_eq!(convert_to_superscript("y=mx+b"), "ʸ⁼ᵐˣ⁺ᵇ"); // all have superscripts
        // Characters without superscript remain unchanged
        assert_eq!(convert_to_superscript("C6H12O6"), "C⁶ᴴ¹²ᴼ⁶"); // 'C' has no superscript
    }

    #[test]
    fn test_convert_to_subscript() {
        // Uppercase 'H' and 'O' have no subscript, only '2' converts
        assert_eq!(convert_to_subscript("H2O"), "H₂O");
        // Lowercase 'h', '2', and 'o' all have subscripts
        assert_eq!(convert_to_subscript("h2o"), "ₕ₂ₒ");
        // 'x' has a subscript equivalent 'ₓ', so it gets converted
        assert_eq!(convert_to_subscript("x0"), "ₓ₀");
        // Characters without subscript remain unchanged
        assert_eq!(convert_to_subscript("abc"), "ₐbc"); // only 'a' has subscript
    }

    #[test]
    fn test_can_convert_checks() {
        assert!(can_convert_to_superscript("123"));
        assert!(can_convert_to_superscript("n+1"));
        assert!(!can_convert_to_superscript("query"));
        
        assert!(can_convert_to_subscript("123"));
        assert!(!can_convert_to_subscript("abc"));
    }
}

