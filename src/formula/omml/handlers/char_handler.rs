// Character element handler

use crate::formula::omml::elements::ElementContext;

/// Handler for character elements (used within properties)
pub struct CharHandler;

impl CharHandler {
    pub fn handle_end<'arena>(
        elem: &[u8],
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if !context.text.is_empty()
            && let Some(parent) = parent_context {
                match std::str::from_utf8(elem).unwrap_or("") {
                    "begChr" | "m:begChr" => {
                        parent.properties.delimiter_open_char = Some(context.text.as_str().to_string());
                    }
                    "endChr" | "m:endChr" => {
                        parent.properties.delimiter_close_char = Some(context.text.as_str().to_string());
                    }
                    _ => {
                        parent.properties.chr = Some(context.text.as_str().to_string());
                    }
                }
            }
    }
}
