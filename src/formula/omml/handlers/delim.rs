// Delimiter element handler

use crate::formula::ast::*;
use crate::formula::omml::attributes::{get_attribute_value, parse_fence_type};
use crate::formula::omml::elements::ElementContext;
use crate::formula::omml::properties::parse_delimiter_properties;
use quick_xml::events::BytesStart;

/// Handler for delimiter (fenced) elements
pub struct DelimiterHandler;

impl DelimiterHandler {
    pub fn handle_start<'arena>(
        elem: &BytesStart,
        context: &mut ElementContext<'arena>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let attrs: Vec<_> = elem.attributes().filter_map(|a| a.ok()).collect();

        // Parse delimiter properties
        context.properties = parse_delimiter_properties(&attrs);

        // Parse fence characters using SIMD-accelerated parsing
        let open_val = get_attribute_value(&attrs, "begChr");
        let close_val = get_attribute_value(&attrs, "endChr");
        let (open_fence, close_fence) = parse_fence_type(open_val.as_deref(), close_val.as_deref());

        if let Some(open) = open_fence {
            context.fence_open = Some(open);
        }
        if let Some(close) = close_fence {
            context.fence_close = Some(close);
        }

        // Parse separator character
        let sep_val = get_attribute_value(&attrs, "sepChr");
        if let Some(sep) = sep_val {
            context.properties.delimiter_separator_char = Some(sep);
        }
    }

    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        arena: &'arena bumpalo::Bump,
    ) {
        // Use fence characters parsed in handle_start, or fall back to properties
        let open = context.fence_open.unwrap_or_else(|| {
            context
                .properties
                .delimiter_open_char
                .as_deref()
                .and_then(|s| parse_fence_type(Some(s), None).0)
                .unwrap_or(Fence::Paren)
        });

        let close = context.fence_close.unwrap_or_else(|| {
            context
                .properties
                .delimiter_close_char
                .as_deref()
                .and_then(|s| parse_fence_type(None, Some(s)).1)
                .unwrap_or(Fence::Paren)
        });

        let content = if context.children.is_empty() {
            Vec::new()
        } else {
            context.children.clone()
        };

        // Use separator from either context properties or element properties
        let separator = context
            .properties
            .delimiter_separator_char
            .as_ref()
            .or(context.properties.delimiter_separator_char.as_ref())
            .map(|s| std::borrow::Cow::Borrowed(arena.alloc_str(s)));

        let node = MathNode::Fenced {
            open,
            content,
            close,
            separator,
        };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}
