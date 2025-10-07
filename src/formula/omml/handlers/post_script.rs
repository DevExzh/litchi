// Post-script element handler

use crate::formula::omml::elements::ElementContext;

/// Handler for post-script elements
pub struct PostScriptHandler;

impl PostScriptHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            // Post-scripts are handled by the superscript/subscript elements
            crate::formula::omml::utils::extend_vec_efficient(&mut parent.post_scripts, context.children.clone());
        }
    }
}
