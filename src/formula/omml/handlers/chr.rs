// Handler for character property element (chr)

use crate::formula::omml::elements::ElementContext;

/// Handler for character property element (chr)
#[allow(dead_code)] // Handler implementation, used via the char_handler module
pub struct ChrHandler;

#[allow(dead_code)] // Handler implementation
impl ChrHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            // Set character property based on element content
            let text_content = context.text.as_str().trim();
            if !text_content.is_empty() {
                parent.properties.chr = Some(text_content.to_string());
            }
        }
    }
}
