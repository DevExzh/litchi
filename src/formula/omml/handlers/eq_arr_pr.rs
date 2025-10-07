// Equation array properties element handler

use crate::formula::omml::elements::ElementContext;

/// Handler for equation array properties elements
pub struct EqArrPrHandler;

impl EqArrPrHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            // Store the parsed properties in the parent context
            parent.properties = context.properties.clone();
            // Pass children up
            crate::formula::omml::utils::extend_vec_efficient(&mut parent.children, context.children.clone());
        }
    }
}
