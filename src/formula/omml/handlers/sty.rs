// Handler for style property element (sty)

use crate::formula::omml::elements::ElementContext;

/// Handler for style property (m:sty)
pub struct StyHandler;

impl StyHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            // Set display style based on element content
            let text_content = context.text.as_str().trim();
            if !text_content.is_empty() {
                parent.properties.display_style = Some(matches!(text_content, "d" | "display" | "1" | "true"));
                parent.properties.run_math_style = Some(text_content.to_string());
            }
        }
    }
}

