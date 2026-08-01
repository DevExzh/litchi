// Handler for style property element (sty)

use crate::omml::elements::ElementContext;

/// Handler for style property (m:sty)
pub struct StyHandler;

impl StyHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: formatting handler, sets flags in context
    ) {
        if let Some(parent) = parent_context {
            // Set display style from the m:val attribute or element content
            let value = context
                .character_data
                .take()
                .unwrap_or_else(|| context.text.as_str().trim().to_string());
            if !value.is_empty() {
                parent.properties.display_style =
                    Some(matches!(value.as_str(), "d" | "display" | "1" | "true"));
                // ECMA-376 ST_Style values (b, i, bi) carry bold/italic style;
                // map them onto the math variant unless m:scr already set one.
                if matches!(value.as_str(), "b" | "i" | "bi")
                    && parent.properties.math_variant.is_none()
                {
                    parent.properties.math_variant = Some(value.clone());
                }
                parent.properties.run_math_style = Some(value);
            }
        }
    }
}
