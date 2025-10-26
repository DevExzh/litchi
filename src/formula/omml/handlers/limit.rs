// Limit element handler

use crate::formula::omml::elements::ElementContext;

/// Handler for limit elements
pub struct LimitHandler;

impl LimitHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: limit elements are owned Vec from context
    ) {
        if let Some(parent) = parent_context {
            // Limits are handled by the specific handlers above
            crate::formula::omml::utils::extend_vec_efficient(
                &mut parent.children,
                context.children.clone(),
            );
        }
    }
}
