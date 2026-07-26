//! AST construct to MTEF 5 template selector/variation mapping
//!
//! MTEF describes structured maths with numbered templates: a selector picks
//! the family (fraction, root, fence, ...) and a variation picks the form within
//! it. This module owns that mapping; [`super::node`] owns the slot contents.
//!
//! The numbering used here is the MTEF 5 table (fences `0..=9`, root `10`,
//! fraction `11`, ...), which differs from the MTEF 1-4 table.

use crate::ast::{Fence, FractionType, LargeOperator, LineStyle, Position};
use crate::mtef::constants::*;

/// A template identified by its selector and variation
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Template {
    /// Template family
    pub selector: u8,
    /// Form within the family
    pub variation: u16,
}

impl Template {
    /// Build a template descriptor
    pub(super) const fn new(selector: u8, variation: u16) -> Self {
        Self {
            selector,
            variation,
        }
    }
}

/// Template selector that draws a given delimiter
///
/// Returns `None` for [`Fence::None`], which is drawn by omitting the template.
fn fence_selector(fence: Fence) -> Option<u8> {
    match fence {
        Fence::Paren => Some(TMPL_PAREN),
        Fence::Bracket => Some(TMPL_BRACK),
        Fence::Brace | Fence::CurlyBrace => Some(TMPL_BRACE),
        Fence::Angle | Fence::AngleBracket => Some(TMPL_ANGLE),
        Fence::Pipe => Some(TMPL_BAR),
        Fence::DoublePipe => Some(TMPL_DBAR),
        Fence::Floor => Some(TMPL_FLOOR),
        Fence::Ceiling => Some(TMPL_CEILING),
        Fence::SquareBracket => Some(TMPL_OBRACK),
        Fence::None => None,
    }
}

/// Glyph nibble used by the mixed-delimiter interval template
///
/// The interval variation stores the opening glyph in its low nibble and the
/// closing glyph in the next one, where each nibble names a concrete glyph
/// rather than a role: `(` and `)` are distinct values.
fn interval_glyph(fence: Fence, opening: bool) -> Option<u16> {
    match (fence, opening) {
        (Fence::Paren, true) => Some(TV_INTERVAL_LPAREN),
        (Fence::Paren, false) => Some(TV_INTERVAL_RPAREN),
        (Fence::Bracket, true) => Some(TV_INTERVAL_LBRACK),
        (Fence::Bracket, false) => Some(TV_INTERVAL_RBRACK),
        _ => None,
    }
}

/// Pick the template that draws `open ... close`
///
/// Returns `None` when neither side is drawn, in which case the content is
/// emitted without a surrounding template.
pub(super) fn fence_template(open: Fence, close: Fence) -> Option<Template> {
    match (open, close) {
        (Fence::None, Fence::None) => None,
        (Fence::None, closing) => fence_selector(closing).map(|s| Template::new(s, TV_FENCE_RIGHT)),
        (opening, Fence::None) => fence_selector(opening).map(|s| Template::new(s, TV_FENCE_LEFT)),
        (opening, closing) if opening == closing => {
            fence_selector(opening).map(|s| Template::new(s, TV_FENCE_BOTH))
        },
        (opening, closing) => {
            // Mismatched delimiters: the interval template can draw any mix of
            // parentheses and square brackets. Anything else keeps the opening
            // delimiter on both sides, which is what MathType offers.
            match (
                interval_glyph(opening, true),
                interval_glyph(closing, false),
            ) {
                (Some(left), Some(right)) => Some(Template::new(
                    TMPL_INTERVAL,
                    left | (right << TV_INTERVAL_CLOSE_SHIFT),
                )),
                _ => fence_selector(opening).map(|s| Template::new(s, TV_FENCE_BOTH)),
            }
        },
    }
}

/// Pick the fraction template for a fraction style
///
/// `FractionType::Skewed` has no MTEF form of its own and shares the slashed
/// (linear) template with `NoBar`.
pub(super) fn fraction_template(frac_type: Option<FractionType>) -> Template {
    let variation = match frac_type {
        Some(FractionType::NoBar) | Some(FractionType::Skewed) => TV_FRACT_SLASH,
        Some(FractionType::Bar) | None => TV_FRACT_BAR,
    };
    Template::new(TMPL_FRACT, variation)
}

/// Pick the root template; `has_index` selects the nth-root form
pub(super) fn root_template(has_index: bool) -> Template {
    Template::new(
        TMPL_ROOT,
        if has_index {
            TV_ROOT_NTH
        } else {
            TV_ROOT_SQUARE
        },
    )
}

/// Pick the over- or underbar template for a rule style
pub(super) fn bar_template(selector: u8, style: Option<LineStyle>) -> Template {
    let variation = match style {
        Some(LineStyle::Double) => TV_BAR_DOUBLE,
        _ => TV_BAR_SINGLE,
    };
    Template::new(selector, variation)
}

/// Pick the horizontal brace template for a group character's position
pub(super) fn group_char_template(position: Option<Position>) -> Template {
    let variation = match position {
        Some(Position::Top) => TV_HBRACE_UPPER,
        _ => TV_HBRACE_LOWER,
    };
    Template::new(TMPL_HBRACE, variation)
}

/// Pick the template that draws a large operator
///
/// Returns `None` for the word-shaped operators (`lim`, `max`, ...), which
/// MathType spells out with function-typeface characters and decorates with the
/// limit template instead.
pub(super) fn large_op_template(operator: LargeOperator, has_limits: bool) -> Option<Template> {
    let template = match operator {
        LargeOperator::Sum => Template::new(TMPL_SUM, TV_DEFAULT),
        LargeOperator::Product => Template::new(TMPL_PROD, TV_DEFAULT),
        LargeOperator::Coproduct => Template::new(TMPL_COPROD, TV_DEFAULT),
        LargeOperator::Union | LargeOperator::BigUnion => Template::new(TMPL_UNION, TV_DEFAULT),
        LargeOperator::Intersection | LargeOperator::BigIntersection => {
            Template::new(TMPL_INTER, TV_DEFAULT)
        },
        LargeOperator::Integral => Template::new(
            TMPL_INTEG,
            if has_limits {
                TV_INTEG_SINGLE_LIMITS
            } else {
                TV_INTEG_SINGLE
            },
        ),
        LargeOperator::DoubleIntegral => Template::new(TMPL_INTEG, TV_INTEG_DOUBLE),
        LargeOperator::TripleIntegral => Template::new(TMPL_INTEG, TV_INTEG_TRIPLE),
        LargeOperator::ContourIntegral => Template::new(TMPL_INTEG, TV_INTEG_CONTOUR),
        LargeOperator::SurfaceIntegral => Template::new(TMPL_INTEG, TV_INTEG_SURFACE),
        LargeOperator::VolumeIntegral => Template::new(TMPL_INTEG, TV_INTEG_VOLUME),
        LargeOperator::Limit
        | LargeOperator::Max
        | LargeOperator::Min
        | LargeOperator::Supremum
        | LargeOperator::Infimum
        | LargeOperator::ArgMax
        | LargeOperator::ArgMin => return None,
    };
    Some(template)
}

/// Spell a word-shaped large operator for the function typeface
pub(super) fn large_op_word(operator: LargeOperator) -> Option<&'static str> {
    match operator {
        LargeOperator::Limit => Some("lim"),
        LargeOperator::Max => Some("max"),
        LargeOperator::Min => Some("min"),
        LargeOperator::Supremum => Some("sup"),
        LargeOperator::Infimum => Some("inf"),
        LargeOperator::ArgMax => Some("argmax"),
        LargeOperator::ArgMin => Some("argmin"),
        _ => None,
    }
}

/// Pick the script template for the scripts that are present
///
/// All three script templates carry two slots (subscript then superscript); the
/// selector states which of them the renderer should draw.
pub(super) fn script_template(has_sub: bool, has_sup: bool) -> Template {
    let selector = match (has_sub, has_sup) {
        (true, true) => TMPL_SUBSUP,
        (true, false) => TMPL_SUB,
        _ => TMPL_SUP,
    };
    Template::new(selector, TV_DEFAULT)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn matching_fences_use_the_both_sides_variation() {
        assert_eq!(
            fence_template(Fence::Paren, Fence::Paren),
            Some(Template::new(TMPL_PAREN, TV_FENCE_BOTH))
        );
    }

    #[test]
    fn one_sided_fences_pick_the_matching_variation() {
        assert_eq!(
            fence_template(Fence::Brace, Fence::None),
            Some(Template::new(TMPL_BRACE, TV_FENCE_LEFT))
        );
        assert_eq!(
            fence_template(Fence::None, Fence::Bracket),
            Some(Template::new(TMPL_BRACK, TV_FENCE_RIGHT))
        );
    }

    #[test]
    fn mixed_delimiters_use_the_interval_template() {
        // "[a, b)" is the LBRP form, variation 18 in the MTEF template table.
        assert_eq!(
            fence_template(Fence::Bracket, Fence::Paren),
            Some(Template::new(TMPL_INTERVAL, 18))
        );
    }

    #[test]
    fn unfenced_content_needs_no_template() {
        assert_eq!(fence_template(Fence::None, Fence::None), None);
    }

    #[test]
    fn word_operators_have_no_template() {
        assert_eq!(large_op_template(LargeOperator::Limit, true), None);
        assert_eq!(large_op_word(LargeOperator::Limit), Some("lim"));
        assert_eq!(large_op_word(LargeOperator::Sum), None);
    }
}
