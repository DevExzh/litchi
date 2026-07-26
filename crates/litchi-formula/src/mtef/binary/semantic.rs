// Semantic classification of the LaTeX spellings the MTEF reader resolves
//
// `charset::lookup_character` answers with a LaTeX spelling such as `\alpha `
// or `\leq `, because the tables were ported from rtf2latex2e, which emitted
// LaTeX directly. The AST has dedicated variants for those glyphs, so this
// module turns a spelling back into the node it names.
//
// The command tables live in `latex::parse::commands` and are shared with the
// LaTeX parser: a spelling and the control sequence a user would type are the
// same string, so both front ends resolve them through one table rather than
// three. Only glyphs the parser has no variant for need the local
// `SPELLING_GLYPHS` table, and those become `Symbol`s carrying their codepoint.
//
// Nothing here fails: a spelling that matches no table yields `None` and the
// caller keeps the raw text, which is what the reader did for every character
// before.

use crate::ast::{MathNode, Symbol};
use crate::latex::parse::commands::{
    DELIMITER_GLYPHS, FUNCTION_WORDS, FUNCTIONS, LARGE_OPERATORS, OPERATORS, PREDEFINED_SYMBOLS,
    SPACES, ascii_operator,
};
use std::borrow::Cow;

/// Single-digit spellings, borrowed so a digit costs no allocation.
const DIGITS: [&str; 10] = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"];

/// Codepoints for spellings that no AST variant names.
///
/// Keys are complete spellings rather than bare command names so that forms
/// like `^\circ` — the degree sign, which MathType spells as a superscript —
/// resolve too. Every entry produces a [`MathNode::Symbol`] carrying the
/// codepoint, which both the LaTeX and the OMML writer render as the glyph.
static SPELLING_GLYPHS: phf::Map<&'static str, char> = phf::phf_map! {
    "\\_"                => '_',
    "\\o"                => 'ø',
    "^\\circ"            => '°',
    "\\ni"               => '∋',
    "\\bot"              => '⊥',
    "\\Im"               => 'ℑ',
    "\\Re"               => 'ℜ',
    "\\wp"               => '℘',
    "\\hbar"             => 'ℏ',
    "\\ell"              => 'ℓ',
    "\\otimes"           => '⊗',
    "\\oplus"            => '⊕',
    "\\nsubset"          => '⊄',
    "\\not\\subset"      => '⊄',
    "\\smallint"         => '∫',
    "\\lozenge"          => '◊',
    "\\backepsilon"      => '϶',
    "\\lambdabar"        => 'ƛ',
    "\\ll"               => '≪',
    "\\gg"               => '≫',
    "\\doteq"            => '≐',
    "\\prec"             => '≺',
    "\\succ"             => '≻',
    "\\vartriangleleft"  => '⊲',
    "\\vartriangleright" => '⊳',
    "\\mapsto"           => '↦',
    "\\longmapsto"       => '⟼',
    "\\hookleftarrow"    => '↩',
    "\\Leftarrow"        => '⇐',
    "\\Rightarrow"       => '⇒',
    "\\Leftrightarrow"   => '⇔',
    "\\Uparrow"          => '⇑',
    "\\Downarrow"        => '⇓',
    "\\Updownarrow"      => '⇕',
};

/// Resolve the LaTeX spelling of one MTEF character into a semantic node.
///
/// Returns `None` when the spelling names nothing the AST models — a plain
/// letter, a punctuation mark or an unrecognised command — so the caller can
/// keep it as text.
pub(super) fn node_for_spelling(spelling: &str) -> Option<MathNode<'static>> {
    let spelling = spelling.trim();
    if let Some(node) = command_node(spelling) {
        return Some(node);
    }
    if let Some((&key, &glyph)) = SPELLING_GLYPHS.get_entry(spelling) {
        return Some(symbol_node(key, glyph));
    }
    literal_node(spelling)
}

/// Resolve a run of characters set in the function typeface into a function.
///
/// The typeface itself marks the run as a function name, so an unknown name
/// still becomes a [`MathNode::Function`] rather than text.
pub(super) fn node_for_function<'a>(name: &str) -> MathNode<'a> {
    if let Some(&function) = FUNCTIONS.get(name) {
        return MathNode::PredefinedFunction {
            function,
            argument: Vec::new(),
        };
    }
    let name = match FUNCTION_WORDS.get_key(name) {
        Some(&known) => Cow::Borrowed(known),
        None => Cow::Owned(name.to_owned()),
    };
    MathNode::Function {
        name,
        argument: Vec::new(),
    }
}

/// Resolve a spelling that is a single control sequence.
fn command_node(spelling: &str) -> Option<MathNode<'static>> {
    let name = spelling
        .strip_prefix('\\')
        .filter(|name| is_command(name))?;

    if let Some(&symbol) = PREDEFINED_SYMBOLS.get(name) {
        return Some(MathNode::PredefinedSymbol(symbol));
    }
    if let Some(&operator) = OPERATORS.get(name) {
        return Some(MathNode::Operator(operator));
    }
    if let Some(&function) = FUNCTIONS.get(name) {
        return Some(MathNode::PredefinedFunction {
            function,
            argument: Vec::new(),
        });
    }
    if let Some(&name) = FUNCTION_WORDS.get_key(name) {
        return Some(MathNode::Function {
            name: Cow::Borrowed(name),
            argument: Vec::new(),
        });
    }
    if let Some(&operator) = LARGE_OPERATORS.get(name) {
        return Some(MathNode::LargeOp {
            operator,
            lower_limit: None,
            upper_limit: None,
            integrand: None,
            hide_lower: true,
            hide_upper: true,
        });
    }
    if let Some(&space) = SPACES.get(name) {
        return Some(MathNode::Space(space));
    }
    if let Some((&name, &glyph)) = DELIMITER_GLYPHS.get_entry(name) {
        return Some(symbol_node(name, glyph));
    }
    None
}

/// Resolve a bare character: a digit is a number, `+` and friends operators.
fn literal_node(spelling: &str) -> Option<MathNode<'static>> {
    let mut chars = spelling.chars();
    let single = chars.next().filter(|_| chars.next().is_none())?;

    if let Some(digit) = single.to_digit(10) {
        return Some(MathNode::Number(Cow::Borrowed(DIGITS[digit as usize])));
    }
    ascii_operator(single).map(MathNode::Operator)
}

/// Build a symbol node, naming it after the spelling that produced it.
fn symbol_node(spelling: &'static str, unicode: char) -> MathNode<'static> {
    MathNode::Symbol(Symbol {
        name: Cow::Borrowed(spelling.trim_start_matches('\\')),
        unicode: Some(unicode),
        variant: None,
    })
}

/// Whether `name` is a control sequence name: letters, or one symbol character.
fn is_command(name: &str) -> bool {
    let mut chars = name.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !first.is_ascii_alphabetic() {
        return chars.next().is_none();
    }
    chars.all(|ch| ch.is_ascii_alphabetic())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ast::{FunctionName, LargeOperator, Operator, PredefinedSymbol, SpaceType};

    #[test]
    fn greek_letters_become_predefined_symbols() {
        assert_eq!(
            node_for_spelling("\\alpha "),
            Some(MathNode::PredefinedSymbol(PredefinedSymbol::Alpha))
        );
        assert_eq!(
            node_for_spelling("\\Delta "),
            Some(MathNode::PredefinedSymbol(PredefinedSymbol::DeltaCap))
        );
        assert_eq!(
            node_for_spelling("\\infty "),
            Some(MathNode::PredefinedSymbol(PredefinedSymbol::Infinity))
        );
    }

    #[test]
    fn relations_become_operators() {
        assert_eq!(
            node_for_spelling("\\leq "),
            Some(MathNode::Operator(Operator::LessThanOrEqual))
        );
        assert_eq!(
            node_for_spelling("\\neq "),
            Some(MathNode::Operator(Operator::NotEquals))
        );
    }

    #[test]
    fn arrows_become_operators() {
        assert_eq!(
            node_for_spelling("\\rightarrow "),
            Some(MathNode::Operator(Operator::RightArrow))
        );
        assert_eq!(
            node_for_spelling("\\leftrightarrow "),
            Some(MathNode::Operator(Operator::LeftRightArrow))
        );
    }

    #[test]
    fn dots_become_operators() {
        assert_eq!(
            node_for_spelling("\\cdots "),
            Some(MathNode::Operator(Operator::CDots))
        );
    }

    #[test]
    fn function_names_become_functions() {
        assert_eq!(
            node_for_spelling("\\sin "),
            Some(MathNode::PredefinedFunction {
                function: FunctionName::Sin,
                argument: Vec::new(),
            })
        );
        assert_eq!(
            node_for_function("cos"),
            MathNode::PredefinedFunction {
                function: FunctionName::Cos,
                argument: Vec::new(),
            }
        );
    }

    #[test]
    fn function_words_without_a_variant_keep_their_name() {
        assert_eq!(
            node_for_function("limsup"),
            MathNode::Function {
                name: Cow::Borrowed("limsup"),
                argument: Vec::new(),
            }
        );
        assert_eq!(
            node_for_function("wobble"),
            MathNode::Function {
                name: Cow::Owned("wobble".to_owned()),
                argument: Vec::new(),
            }
        );
    }

    #[test]
    fn large_operators_keep_their_variant() {
        assert!(matches!(
            node_for_spelling("\\sum "),
            Some(MathNode::LargeOp {
                operator: LargeOperator::Sum,
                ..
            })
        ));
    }

    #[test]
    fn spacing_commands_become_spaces() {
        assert_eq!(
            node_for_spelling("\\quad "),
            Some(MathNode::Space(SpaceType::Quad))
        );
        assert_eq!(
            node_for_spelling("\\,"),
            Some(MathNode::Space(SpaceType::Thin))
        );
    }

    #[test]
    fn glyphs_outside_the_command_tables_become_symbols() {
        let node = node_for_spelling("\\otimes ");
        assert_eq!(
            node,
            Some(MathNode::Symbol(Symbol {
                name: Cow::Borrowed("otimes"),
                unicode: Some('⊗'),
                variant: None,
            }))
        );
        assert!(matches!(
            node_for_spelling("^\\circ "),
            Some(MathNode::Symbol(Symbol {
                unicode: Some('°'),
                ..
            }))
        ));
    }

    #[test]
    fn delimiters_become_symbols() {
        assert!(matches!(
            node_for_spelling("\\langle "),
            Some(MathNode::Symbol(Symbol {
                unicode: Some('⟨'),
                ..
            }))
        ));
    }

    #[test]
    fn plain_letters_are_not_classified() {
        assert_eq!(node_for_spelling("A"), None);
        assert_eq!(node_for_spelling("x"), None);
        assert_eq!(node_for_spelling("xy"), None);
    }

    #[test]
    fn digits_become_numbers() {
        assert_eq!(
            node_for_spelling("7"),
            Some(MathNode::Number(Cow::Borrowed("7")))
        );
    }

    #[test]
    fn ascii_operators_are_recognised() {
        assert_eq!(
            node_for_spelling("="),
            Some(MathNode::Operator(Operator::Equals))
        );
        assert_eq!(
            node_for_spelling("+"),
            Some(MathNode::Operator(Operator::Plus))
        );
        // `*` renders as `\ast`, so keeping it literal stays lossless.
        assert_eq!(node_for_spelling("*"), None);
    }

    #[test]
    fn unknown_spellings_degrade_to_nothing() {
        assert_eq!(node_for_spelling("\\notacommand "), None);
        assert_eq!(node_for_spelling("\\mathop{\\rm var} "), None);
        assert_eq!(node_for_spelling("{}"), None);
        assert_eq!(node_for_spelling(""), None);
    }
}
