// Handler for position property element (pos)

use crate::omml::elements::ElementContext;

/// Handler for position property element (pos)
pub struct PosHandler;

impl PosHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: formatting handler, sets flags in context
    ) {
        if let Some(parent) = parent_context {
            // Set position from the m:val attribute or element content
            let value = context
                .character_data
                .take()
                .unwrap_or_else(|| context.text.as_str().trim().to_string());
            if !value.is_empty() {
                parent.properties.accent_position = Some(value);
            }
        }
    }
}
