// Run properties element handler

use crate::formula::omml::elements::ElementContext;

/// Handler for run properties elements (rPr)
pub struct RunPropsHandler;

impl RunPropsHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        // Run properties are stored in the parent context
        if let Some(parent) = parent_context {
            // Copy all run-related properties from this context to parent
            parent.properties.run_literal = context.properties.run_literal;
            parent.properties.run_normal_text = context.properties.run_normal_text.clone();
            parent.properties.run_math_style = context.properties.run_math_style.clone();
            parent.properties.math_variant = context.properties.math_variant.clone();
            parent.properties.display_style = context.properties.display_style;
            parent.properties.font = context.properties.font.clone();
            parent.properties.color = context.properties.color.clone();
            parent.properties.underline = context.properties.underline.clone();
            parent.properties.overline = context.properties.overline.clone();
            parent.properties.strike_through = context.properties.strike_through;
            parent.properties.double_strike_through = context.properties.double_strike_through;
        }
    }
}
