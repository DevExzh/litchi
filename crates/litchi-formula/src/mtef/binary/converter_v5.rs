//! MTEF 5 template conversion
//!
//! MTEF 5 renumbered the template table that MTEF 1-4 used: what version 3 calls
//! selector 14 (fraction) is selector 11 in version 5, and the script templates
//! moved from a single selector with three variations to three selectors. The
//! conversion in [`super::converter`] follows the older numbering, so version 5
//! records are routed here instead.
//!
//! Template slots arrive as LINE records in the order the template table
//! defines: numerator before denominator, radicand before degree, integrand
//! before limits, and subscript before superscript. Script templates carry no
//! base at all — in MTEF the base is the object that precedes the record.
//!
//! Reference: rtf2latex2e `Profile_TEMPLATES_5` and
//! <http://rtf2latex2e.sourceforge.net/MTEF5.html>

use super::objects::*;
use crate::ast::{FractionType, LargeOperator, MathNode, Position, VerticalAlignment};
use crate::mtef::MtefError;
use crate::mtef::constants::*;
use std::borrow::Cow;

/// Character MathType draws for a brace below a group
const GROUP_CHAR_UNDER_BRACE: &str = "\u{23DF}";
/// Character MathType draws for a brace above a group
const GROUP_CHAR_OVER_BRACE: &str = "\u{23DE}";

/// Whether a template's base is the object preceding it in the line
///
/// The script templates hold only their scripts; everything else carries its
/// base in a slot.
pub(super) fn takes_preceding_base(selector: u8) -> bool {
    matches!(selector, TMPL_SUB | TMPL_SUP | TMPL_SUBSUP)
}

impl<'arena> super::parser::MtefBinaryParser<'arena> {
    /// Convert an MTEF 5 template record into an AST node
    ///
    /// `base` holds the object that preceded the record, which only the script
    /// templates consume.
    pub(super) fn convert_template_v5(
        &self,
        tmpl_obj: &MtefTemplate,
        base: Vec<MathNode<'arena>>,
    ) -> Result<MathNode<'arena>, MtefError> {
        let mut slots = self.template_slots_v5(tmpl_obj)?;

        match tmpl_obj.selector {
            TMPL_ANGLE..=TMPL_INTERVAL => self.convert_fence_template(tmpl_obj),
            TMPL_ROOT => {
                let index = take_slot(&mut slots, 1);
                Ok(MathNode::Root {
                    base: take_slot(&mut slots, 0),
                    index: if index.is_empty() { None } else { Some(index) },
                })
            },
            TMPL_FRACT => Ok(MathNode::Frac {
                numerator: take_slot(&mut slots, 0),
                denominator: take_slot(&mut slots, 1),
                line_thickness: None,
                frac_type: fraction_type(tmpl_obj.variation),
            }),
            TMPL_UBAR | TMPL_OBAR => self.convert_decoration_template(tmpl_obj),
            TMPL_ARROW => self.convert_arrow_template(tmpl_obj),
            TMPL_INTEG | TMPL_INTOP | TMPL_SUM | TMPL_SUMOP | TMPL_PROD | TMPL_COPROD
            | TMPL_UNION | TMPL_INTER => {
                Ok(large_op(tmpl_obj.selector, tmpl_obj.variation, &mut slots))
            },
            TMPL_LIM => Ok(limit_node(&mut slots)),
            TMPL_HBRACE | TMPL_HBRACK => Ok(group_char(tmpl_obj.variation, &mut slots)),
            TMPL_SUB | TMPL_SUP | TMPL_SUBSUP => {
                Ok(script_node(tmpl_obj.selector, base, &mut slots))
            },
            _ => Ok(MathNode::Row(slots.into_iter().flatten().collect())),
        }
    }

    /// Collect a template's slots, one entry per LINE record
    ///
    /// Empty slots are preserved so that positional lookups stay correct: a
    /// superscript template writes an empty subscript slot ahead of the script
    /// it actually carries.
    fn template_slots_v5(
        &self,
        tmpl_obj: &MtefTemplate,
    ) -> Result<Vec<Vec<MathNode<'arena>>>, MtefError> {
        let mut slots = Vec::new();
        let mut current = tmpl_obj.subobject_list.as_deref();

        while let Some(obj) = current {
            if obj.tag == MtefRecordType::Line
                && let Some(line_obj) = obj.obj_ptr.as_any().downcast_ref::<MtefLine>()
            {
                slots.push(self.convert_line_to_nodes(line_obj)?.unwrap_or_default());
            }
            current = obj.next.as_deref();
        }

        Ok(slots)
    }
}

/// Take a slot by position, leaving an empty slot behind
fn take_slot<'a>(slots: &mut [Vec<MathNode<'a>>], index: usize) -> Vec<MathNode<'a>> {
    slots.get_mut(index).map(std::mem::take).unwrap_or_default()
}

/// Map a fraction variation onto a fraction style
///
/// The slashed variations describe a linear fraction; the remaining ones all
/// draw a horizontal bar, which is the AST's default.
fn fraction_type(variation: u16) -> Option<FractionType> {
    match variation {
        TV_FRACT_BAR | TV_FRACT_SMALL_BAR => None,
        _ => Some(FractionType::NoBar),
    }
}

/// Build a large-operator node from an operator template and its slots
///
/// Slots arrive as integrand, lower limit, upper limit. A limit that MathType
/// does not draw is stored as an empty slot, which becomes a hidden limit.
fn large_op<'a>(selector: u8, variation: u16, slots: &mut [Vec<MathNode<'a>>]) -> MathNode<'a> {
    let operator = match selector {
        TMPL_SUM | TMPL_SUMOP => LargeOperator::Sum,
        TMPL_PROD => LargeOperator::Product,
        TMPL_COPROD => LargeOperator::Coproduct,
        TMPL_UNION => LargeOperator::Union,
        TMPL_INTER => LargeOperator::Intersection,
        _ => match variation {
            TV_INTEG_DOUBLE => LargeOperator::DoubleIntegral,
            TV_INTEG_TRIPLE => LargeOperator::TripleIntegral,
            TV_INTEG_CONTOUR => LargeOperator::ContourIntegral,
            TV_INTEG_SURFACE => LargeOperator::SurfaceIntegral,
            TV_INTEG_VOLUME => LargeOperator::VolumeIntegral,
            _ => LargeOperator::Integral,
        },
    };

    let integrand = take_slot(slots, 0);
    let lower = take_slot(slots, 1);
    let upper = take_slot(slots, 2);

    MathNode::LargeOp {
        operator,
        hide_lower: lower.is_empty(),
        hide_upper: upper.is_empty(),
        lower_limit: (!lower.is_empty()).then_some(lower),
        upper_limit: (!upper.is_empty()).then_some(upper),
        integrand: (!integrand.is_empty()).then_some(integrand),
    }
}

/// Build an under/over node from a limit template's slots
///
/// Slots arrive as base, lower and upper material.
fn limit_node<'a>(slots: &mut [Vec<MathNode<'a>>]) -> MathNode<'a> {
    let base = take_slot(slots, 0);
    let under = take_slot(slots, 1);
    let over = take_slot(slots, 2);

    match (under.is_empty(), over.is_empty()) {
        (false, false) => MathNode::UnderOver {
            base,
            under,
            over,
            position: None,
        },
        (false, true) => MathNode::Under {
            base,
            under,
            position: None,
        },
        (true, false) => MathNode::Over {
            base,
            over,
            position: None,
        },
        (true, true) => MathNode::Row(base),
    }
}

/// Build a group-character node from a horizontal brace template
fn group_char<'a>(variation: u16, slots: &mut [Vec<MathNode<'a>>]) -> MathNode<'a> {
    let upper = variation == TV_HBRACE_UPPER;
    MathNode::GroupChar {
        base: Box::new(take_slot(slots, 0)),
        character: Some(Cow::Borrowed(if upper {
            GROUP_CHAR_OVER_BRACE
        } else {
            GROUP_CHAR_UNDER_BRACE
        })),
        position: Some(if upper {
            Position::Top
        } else {
            Position::Bottom
        }),
        vertical_alignment: Some(VerticalAlignment::Center),
    }
}

/// Build a script node from a script template's slots and its base
///
/// All three script templates carry a subscript slot followed by a superscript
/// slot; the selector states which of them is drawn. Streams that write only the
/// slot they use are handled by falling back to the first non-empty slot.
fn script_node<'a>(
    selector: u8,
    base: Vec<MathNode<'a>>,
    slots: &mut [Vec<MathNode<'a>>],
) -> MathNode<'a> {
    let subscript = take_slot(slots, 0);
    let superscript = take_slot(slots, 1);

    match selector {
        TMPL_SUB => MathNode::Sub { base, subscript },
        TMPL_SUP => MathNode::Power {
            base,
            exponent: if superscript.is_empty() {
                subscript
            } else {
                superscript
            },
        },
        _ => MathNode::SubSup {
            base,
            subscript,
            superscript,
        },
    }
}
