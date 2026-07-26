// Handler for script/style property element (scr)

use crate::omml::elements::ElementContext;

/// Handler for script/style property (m:scr)
pub struct ScrHandler;

impl ScrHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: formatting handler, sets flags in context
    ) {
        if let Some(parent) = parent_context {
            // Set math variant from the m:val attribute or element content
            let value = context.character_data.take().unwrap_or_else(|| {
                context.text.as_str().trim().to_string()
            });
            if !value.is_empty() {
                parent.properties.math_variant = Some(value.clone());
                parent.properties.run_math_style = Some(value);
            }
        }
    }
}
