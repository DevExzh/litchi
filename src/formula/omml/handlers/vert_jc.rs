// Handler for vertical justification property element (vertJc)

use crate::formula::omml::elements::ElementContext;

/// Handler for vertical justification property element (vertJc)
pub struct VertJcHandler;

impl VertJcHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: formatting handler, sets flags in context
    ) {
        if let Some(parent) = parent_context {
            // Set vertical alignment property based on element content
            let text_content = context.text.as_str().trim();
            if !text_content.is_empty() {
                parent.properties.vertical_alignment = Some(text_content.to_string());
            }
        }
    }
}
