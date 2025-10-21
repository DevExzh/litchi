// Handler for literal text property element (lit)

use crate::formula::omml::elements::ElementContext;

/// Handler for literal text property (m:lit)
pub struct LitHandler;

impl LitHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            // Set literal property based on element content
            let text_content = context.text.as_str().trim();
            if !text_content.is_empty() {
                parent.properties.run_literal =
                    Some(text_content == "1" || text_content.to_lowercase() == "true");
            }
        }
    }
}
