// Handler for literal text property element (lit)

use crate::omml::elements::ElementContext;

/// Handler for literal text property (m:lit)
pub struct LitHandler;

impl LitHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: formatting handler, sets flags in context
    ) {
        if let Some(parent) = parent_context {
            // Set literal property from the m:val attribute or element content.
            // A bare `<m:lit/>` element means "on" (CT_OnOff semantics).
            let value = context.character_data.take().or_else(|| {
                let text_content = context.text.as_str().trim();
                if text_content.is_empty() {
                    None
                } else {
                    Some(text_content.to_string())
                }
            });
            parent.properties.run_literal = Some(match value {
                Some(v) => v == "1" || v.eq_ignore_ascii_case("true") || v == "on",
                None => true,
            });
        }
    }
}
