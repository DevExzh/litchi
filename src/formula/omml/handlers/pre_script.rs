// Pre-script element handler

use crate::formula::omml::elements::ElementContext;

/// Handler for pre-script elements
pub struct PreScriptHandler;

impl PreScriptHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: script positioning, no allocation needed
    ) {
        if let Some(parent) = parent_context {
            // Pre-scripts are handled by the superscript/subscript elements
            crate::formula::omml::utils::extend_vec_efficient(
                &mut parent.pre_scripts,
                context.children.clone(),
            );
        }
    }
}
