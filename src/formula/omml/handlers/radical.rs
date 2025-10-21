// Radical element handler

use crate::formula::ast::*;
use crate::formula::omml::elements::ElementContext;

/// Handler for radical (root) elements
pub struct RadicalHandler;

impl RadicalHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let base = context.base.take().unwrap_or_default();
        let index = context.degree.take();

        let node = MathNode::Root { base, index };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}
