// Script element handlers

use crate::formula::ast::*;
use crate::formula::omml::elements::ElementContext;

/// Handler for superscript elements
pub struct SuperscriptHandler;

impl SuperscriptHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let base = context.base.take().unwrap_or_default();
        let exponent = context.superscript.take().unwrap_or_default();

        let node = MathNode::Power { base, exponent };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}

/// Handler for subscript elements
pub struct SubscriptHandler;

impl SubscriptHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let base = context.base.take().unwrap_or_default();
        let subscript = context.subscript.take().unwrap_or_default();

        let node = MathNode::Sub { base, subscript };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}

/// Handler for subscript-superscript elements
pub struct SubSupHandler;

impl SubSupHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let base = context.base.take().unwrap_or_default();
        let subscript = context.subscript.take().unwrap_or_default();
        let superscript = context.superscript.take().unwrap_or_default();

        let node = MathNode::SubSup {
            base,
            subscript,
            superscript,
        };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}

/// Handler for superscript elements
pub struct SuperscriptElementHandler;

impl SuperscriptElementHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            match parent.element_type {
                crate::formula::omml::elements::ElementType::Superscript
                | crate::formula::omml::elements::ElementType::SubSup => {
                    parent.superscript = Some(context.children.clone());
                },
                crate::formula::omml::elements::ElementType::Nary => {
                    parent.upper_limit = Some(context.children.clone());
                },
                _ => {
                    // Pass children up for other contexts
                    crate::formula::omml::utils::extend_vec_efficient(
                        &mut parent.children,
                        context.children.clone(),
                    );
                },
            }
        }
    }
}

/// Handler for subscript elements
pub struct SubscriptElementHandler;

impl SubscriptElementHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            match parent.element_type {
                crate::formula::omml::elements::ElementType::Subscript
                | crate::formula::omml::elements::ElementType::SubSup => {
                    parent.subscript = Some(context.children.clone());
                },
                crate::formula::omml::elements::ElementType::Nary => {
                    parent.lower_limit = Some(context.children.clone());
                },
                _ => {
                    // Pass children up for other contexts
                    crate::formula::omml::utils::extend_vec_efficient(
                        &mut parent.children,
                        context.children.clone(),
                    );
                },
            }
        }
    }
}
