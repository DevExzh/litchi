// Character element handler

use crate::formula::omml::elements::ElementContext;
use crate::formula::omml::attributes::get_attribute_value;

/// Handler for character elements (used within properties)
pub struct CharHandler;

impl CharHandler {
    pub fn handle_end<'arena>(
        elem: &[u8],
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            // Get character value from either val attribute or text content
            let char_value = get_attribute_value(&context.attributes, "val")
                .or_else(|| if !context.text.is_empty() {
                    Some(context.text.as_str().to_string())
                } else {
                    None
                });

            if let Some(value) = char_value {
                match std::str::from_utf8(elem).unwrap_or("") {
                    "begChr" | "m:begChr" => {
                        parent.properties.delimiter_open_char = Some(value);
                    }
                    "endChr" | "m:endChr" => {
                        parent.properties.delimiter_close_char = Some(value);
                    }
                    _ => {
                        parent.properties.chr = Some(value);
                    }
                }
            }
        }
    }
}
