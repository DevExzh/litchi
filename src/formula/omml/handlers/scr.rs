// Handler for script/style property element (scr)

use crate::formula::omml::elements::ElementContext;

/// Handler for script/style property (m:scr)
pub struct ScrHandler;

impl ScrHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            // Set math variant based on element content
            let text_content = context.text.as_str().trim();
            if !text_content.is_empty() {
                parent.properties.math_variant = Some(text_content.to_string());
                parent.properties.run_math_style = Some(text_content.to_string());
            }
        }
    }
}

