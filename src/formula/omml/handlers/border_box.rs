// Border box element handler

use crate::formula::ast::*;
use crate::formula::omml::elements::ElementContext;

/// Handler for border box elements
pub struct BorderBoxHandler;

impl BorderBoxHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: simple wrapper, children are owned Vec
    ) {
        let content = if context.children.is_empty() {
            Vec::new()
        } else {
            context.children.clone()
        };

        // Border box element - treated as fenced with no fence
        let node = MathNode::Fenced {
            open: Fence::None,
            content,
            close: Fence::None,
            separator: None,
        };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}
