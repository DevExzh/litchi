// Character element handler

use crate::omml::elements::ElementContext;

/// Handler for character elements (used within properties)
pub struct CharHandler;

impl CharHandler {
    pub fn handle_end<'arena>(
        elem: &[u8],
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: character value stored in context properties
    ) {
        if let Some(parent) = parent_context {
            // Get character value from either the m:val attribute (captured
            // during element start) or the element text content.
            let char_value = context.character_data.take().or_else(|| {
                if !context.text.is_empty() {
                    Some(context.text.as_str().to_string())
                } else {
                    None
                }
            });

            if let Some(value) = char_value {
                match std::str::from_utf8(elem).unwrap_or("") {
                    "begChr" | "m:begChr" => {
                        parent.properties.delimiter_open_char = Some(value);
                    },
                    "endChr" | "m:endChr" => {
                        parent.properties.delimiter_close_char = Some(value);
                    },
                    "sepChr" | "m:sepChr" => {
                        parent.properties.delimiter_separator_char = Some(value);
                    },
                    _ => {
                        parent.properties.chr = Some(value);
                    },
                }
            }
        }
    }
}
