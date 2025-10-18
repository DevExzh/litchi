// Handler for position property element (pos)

use crate::formula::omml::elements::ElementContext;

/// Handler for position property element (pos)
pub struct PosHandler;

impl PosHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            // Set position property based on element content
            let text_content = context.text.as_str().trim();
            if !text_content.is_empty() {
                parent.properties.accent_position = Some(text_content.to_string());
            }
        }
    }
}

