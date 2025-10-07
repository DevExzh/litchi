// Spacing element handler

use crate::formula::ast::*;
use crate::formula::omml::elements::ElementContext;
use crate::formula::omml::attributes::{get_attribute_value, parse_space_type};
use crate::formula::omml::properties::parse_spacing_properties;
use quick_xml::events::BytesStart;

/// Handler for spacing elements
pub struct SpacingHandler;

impl SpacingHandler {
    pub fn handle_start<'arena>(
        elem: &BytesStart,
        context: &mut ElementContext<'arena>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let attrs: Vec<_> = elem.attributes().filter_map(|a| a.ok()).collect();

        // Parse spacing properties
        context.properties = parse_spacing_properties(&attrs);

        // Parse spacing value using SIMD-accelerated parsing
        let spacing_val = get_attribute_value(&attrs, "val");
        if let Some(val) = spacing_val {
            context.properties.spacing = Some(val);
        }
    }

    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        // Create a spacing node if spacing is specified
        if let Some(spacing) = &context.properties.spacing {
            // Use the dedicated space type parsing function
            let space_type = parse_space_type(Some(spacing))
                .unwrap_or(SpaceType::Thin);

            let node = MathNode::Space(space_type);

            if let Some(parent) = parent_context {
                parent.children.push(node);
            }
        }
    }
}
