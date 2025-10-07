// Component element handlers

use crate::formula::omml::elements::{ElementContext, ElementType};
use crate::formula::omml::utils::extend_vec_efficient;

/// Handler for numerator elements
pub struct NumeratorHandler;

impl NumeratorHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            if parent.element_type == ElementType::Fraction {
                parent.numerator = Some(context.children.clone());
            } else {
                // Pass children up if not in a fraction context
                extend_vec_efficient(&mut parent.children, context.children.clone());
            }
        }
    }
}

/// Handler for denominator elements
pub struct DenominatorHandler;

impl DenominatorHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            if parent.element_type == ElementType::Fraction {
                parent.denominator = Some(context.children.clone());
            } else {
                // Pass children up if not in a fraction context
                extend_vec_efficient(&mut parent.children, context.children.clone());
            }
        }
    }
}

/// Handler for degree elements (for radicals)
pub struct DegreeHandler;

impl DegreeHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            if parent.element_type == ElementType::Radical {
                parent.degree = Some(context.children.clone());
            } else {
                // Pass children up if not in a radical context
                extend_vec_efficient(&mut parent.children, context.children.clone());
            }
        }
    }
}

/// Handler for base elements
pub struct BaseHandler;

impl BaseHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            match parent.element_type {
                ElementType::Superscript | ElementType::Subscript | ElementType::SubSup => {
                    parent.base = Some(context.children.clone());
                }
                ElementType::Radical => {
                    parent.base = Some(context.children.clone());
                }
                ElementType::Accent | ElementType::Bar | ElementType::GroupChar => {
                    parent.base = Some(context.children.clone());
                }
                _ => {
                    // Pass children up for other contexts
                    extend_vec_efficient(&mut parent.children, context.children.clone());
                }
            }
        }
    }
}

/// Handler for lower limit elements
pub struct LowerLimitHandler;

impl LowerLimitHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            if parent.element_type == ElementType::Nary {
                parent.lower_limit = Some(context.children.clone());
            } else {
                // Pass children up if not in a nary context
                extend_vec_efficient(&mut parent.children, context.children.clone());
            }
        }
    }
}

/// Handler for upper limit elements
pub struct UpperLimitHandler;

impl UpperLimitHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            if parent.element_type == ElementType::Nary {
                parent.upper_limit = Some(context.children.clone());
            } else {
                // Pass children up if not in a nary context
                extend_vec_efficient(&mut parent.children, context.children.clone());
            }
        }
    }
}

/// Handler for integrand elements
pub struct IntegrandHandler;

impl IntegrandHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            if parent.element_type == ElementType::Nary {
                parent.integrand = Some(context.children.clone());
            } else {
                // Pass children up if not in a nary context
                crate::formula::omml::utils::extend_vec_efficient(&mut parent.children, context.children.clone());
            }
        }
    }
}

/// Handler for upper limit elements (limUpp)
pub struct LimUppHandler;

impl LimUppHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            if parent.element_type == ElementType::Nary {
                parent.upper_limit = Some(context.children.clone());
            } else {
                // If not in nary context, treat as overset
                crate::formula::omml::utils::extend_vec_efficient(&mut parent.children, context.children.clone());
            }
        }
    }
}

/// Handler for lower limit elements (limLow)
pub struct LimLowHandler;

impl LimLowHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            if parent.element_type == ElementType::Nary {
                parent.lower_limit = Some(context.children.clone());
            } else {
                // If not in nary context, treat as underset
                crate::formula::omml::utils::extend_vec_efficient(&mut parent.children, context.children.clone());
            }
        }
    }
}
