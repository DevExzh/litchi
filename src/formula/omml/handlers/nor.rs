// Handler for normal text property element (nor)

use crate::formula::omml::elements::ElementContext;

/// Handler for normal text property (m:nor)
pub struct NorHandler;

impl NorHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            // Set normal text font based on element content
            let text_content = context.text.as_str().trim();
            if !text_content.is_empty() {
                parent.properties.font = Some(text_content.to_string());
                parent.properties.run_normal_text = Some(text_content.to_string());
            }
        }
    }
}
