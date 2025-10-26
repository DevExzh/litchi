// Handler for group character properties (groupChrPr)

use crate::formula::omml::elements::ElementContext;

/// Handler for group character properties (groupChrPr)
pub struct GroupChrPrHandler;

impl GroupChrPrHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: property handler, only parses attributes into context
    ) {
        // Group character properties are stored in the parent context
        if let Some(parent) = parent_context {
            // Copy properties from this context to parent
            parent.properties.chr = context.properties.chr.clone();
            parent.properties.accent_position = context.properties.accent_position.clone();
            parent.properties.vertical_alignment = context.properties.vertical_alignment.clone();
        }
    }
}
