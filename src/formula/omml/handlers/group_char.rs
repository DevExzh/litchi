// Group character element handler

use crate::formula::ast::*;
use crate::formula::omml::attributes::{
    get_attribute_value, parse_position_type, parse_vertical_alignment,
};
use crate::formula::omml::elements::ElementContext;
use crate::formula::omml::properties::parse_group_char_properties;
use quick_xml::events::BytesStart;
use std::borrow::Cow;

/// Handler for group character elements
pub struct GroupCharHandler;

impl GroupCharHandler {
    pub fn handle_start<'arena>(
        elem: &BytesStart,
        context: &mut ElementContext<'arena>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let attrs: Vec<_> = elem.attributes().filter_map(|a| a.ok()).collect();

        // Parse group character properties
        context.properties = parse_group_char_properties(&attrs);

        // Parse character using SIMD-accelerated parsing
        let chr_val = get_attribute_value(&attrs, "chr");
        if let Some(chr) = chr_val {
            context.properties.chr = Some(chr);
        }

        // Parse position using dedicated parsing function
        let pos_val = get_attribute_value(&attrs, "pos");
        if let Some(pos) = pos_val {
            context.properties.accent_position = Some(pos);
        }

        // Parse vertical alignment using dedicated parsing function
        let vert_jc_val = get_attribute_value(&attrs, "vertJc");
        if let Some(vert_jc) = vert_jc_val {
            context.properties.vertical_alignment = Some(vert_jc);
        }
    }

    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        arena: &'arena bumpalo::Bump,
    ) {
        let base = if let Some(ref base_content) = context.base {
            base_content.clone()
        } else if context.children.is_empty() {
            Vec::new()
        } else {
            context.children.clone()
        };

        let character = context
            .properties
            .chr
            .as_ref()
            .map(|s| Cow::Borrowed(arena.alloc_str(s)));

        // Use dedicated position parsing function
        let position = context
            .properties
            .accent_position
            .as_ref()
            .and_then(|s| parse_position_type(Some(s)));

        // Use dedicated vertical alignment parsing function
        let vertical_alignment = context
            .properties
            .vertical_alignment
            .as_ref()
            .and_then(|s| parse_vertical_alignment(Some(s)));

        let node = MathNode::GroupChar {
            base: Box::new(base),
            character,
            position,
            vertical_alignment,
        };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}
