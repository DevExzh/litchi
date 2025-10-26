// Phantom element handler

use crate::formula::ast::*;
use crate::formula::omml::elements::ElementContext;

/// Handler for phantom elements
pub struct PhantomHandler;

impl PhantomHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: simple wrapper, children are owned Vec
    ) {
        let content = if context.children.is_empty() {
            Box::new(Vec::new())
        } else {
            Box::new(context.children.clone())
        };

        let node = MathNode::Phantom(content);

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}
