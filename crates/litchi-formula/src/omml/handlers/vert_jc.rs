// Handler for vertical justification property element (vertJc)

use crate::omml::elements::ElementContext;

/// Handler for vertical justification property element (vertJc)
pub struct VertJcHandler;

impl VertJcHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: formatting handler, sets flags in context
    ) {
        if let Some(parent) = parent_context {
            // Set vertical alignment from the m:val attribute or element content
            let value = context.character_data.take().unwrap_or_else(|| {
                context.text.as_str().trim().to_string()
            });
            if !value.is_empty() {
                parent.properties.vertical_alignment = Some(value);
            }
        }
    }
}
