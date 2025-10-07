// Border box element handler

use crate::formula::ast::*;
use crate::formula::omml::elements::ElementContext;

/// Handler for border box elements
pub struct BorderBoxHandler;

impl BorderBoxHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let content = if context.children.is_empty() {
            Vec::new()
        } else {
            context.children.clone()
        };

        // Create border box style from properties
        let style = if context.properties.border_hide_top.is_some()
            || context.properties.border_hide_bottom.is_some()
            || context.properties.border_hide_left.is_some()
            || context.properties.border_hide_right.is_some()
            || context.properties.border_strike_horizontal.is_some()
            || context.properties.border_strike_vertical.is_some()
            || context.properties.border_strike_bltr.is_some()
            || context.properties.border_strike_tlbr.is_some() {
            Some(BorderBoxStyle {
                hide_top: context.properties.border_hide_top.unwrap_or(false),
                hide_bottom: context.properties.border_hide_bottom.unwrap_or(false),
                hide_left: context.properties.border_hide_left.unwrap_or(false),
                hide_right: context.properties.border_hide_right.unwrap_or(false),
                strike_horizontal: context.properties.border_strike_horizontal.unwrap_or(false),
                strike_vertical: context.properties.border_strike_vertical.unwrap_or(false),
                strike_bltr: context.properties.border_strike_bltr.unwrap_or(false),
                strike_tlbr: context.properties.border_strike_tlbr.unwrap_or(false),
            })
        } else {
            None
        };

        let node = MathNode::BorderBox {
            content: Box::new(content),
            style,
        };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}
