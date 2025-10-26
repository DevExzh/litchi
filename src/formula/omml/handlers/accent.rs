// Accent element handler

use crate::formula::ast::*;
use crate::formula::omml::attributes::{
    get_attribute_value, parse_accent_type, parse_position_type,
};
use crate::formula::omml::elements::ElementContext;
use crate::formula::omml::properties::parse_accent_properties;
use quick_xml::events::BytesStart;

/// Handler for accent elements
pub struct AccentHandler;

impl AccentHandler {
    pub fn handle_start<'arena>(
        elem: &BytesStart,
        context: &mut ElementContext<'arena>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let attrs: Vec<_> = elem.attributes().filter_map(|a| a.ok()).collect();

        // Parse accent position from pos attribute
        let pos_val = get_attribute_value(&attrs, "pos");
        if let Some(pos_str) = pos_val {
            context.properties.accent_position = Some(pos_str);
        }

        // Parse accent properties (though chr is now handled as child element)
        context.properties = parse_accent_properties(&attrs);
    }

    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        // Try to get accent type from properties if not set by attribute
        let accent_type = context.accent_type.or_else(|| {
            context
                .properties
                .chr
                .as_deref()
                .and_then(|s| parse_accent_type(Some(s)))
        });

        // Use the detected accent type, or default to Bar (overline) if not recognized
        // This matches the behavior of LibreOffice and plurimath
        let accent_type = accent_type.unwrap_or(AccentType::Bar);

        let base = context
            .base
            .clone()
            .unwrap_or_else(|| context.children.clone());

        // Parse position using the dedicated position parsing function
        let position = context
            .properties
            .accent_position
            .as_ref()
            .and_then(|s| parse_position_type(Some(s)));

        let node = MathNode::Accent {
            base: Box::new(base),
            accent: accent_type,
            position,
        };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}
