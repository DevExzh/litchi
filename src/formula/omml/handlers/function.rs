// Function element handlers

use crate::formula::ast::*;
use crate::formula::omml::elements::ElementContext;
use std::borrow::Cow;

/// Handler for function elements
pub struct FunctionHandler;

impl FunctionHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        arena: &'arena bumpalo::Bump,
    ) {
        // Try to get function name from context, otherwise use generic function
        let argument = if context.children.is_empty() {
            Vec::new()
        } else {
            context.children.clone()
        };

        let name = context.function_name.take()
            .unwrap_or_else(|| "f".to_string());
        let name = arena.alloc_str(&name);

        let node = MathNode::Function {
            name: Cow::Borrowed(name),
            argument,
        };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}

/// Handler for function name elements
pub struct FunctionNameHandler;

impl FunctionNameHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        // Collect text from children to form function name
        let mut name = String::new();
        for child in &context.children {
            if let MathNode::Text(text) = child {
                name.push_str(text.as_ref());
            }
        }

        if !name.is_empty()
            && let Some(parent) = parent_context {
                parent.function_name = Some(name);
            }
    }
}
