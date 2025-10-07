// N-ary operator element handler

use crate::formula::ast::*;
use crate::formula::omml::elements::ElementContext;
use crate::formula::omml::attributes::{get_attribute_value, parse_large_operator};
use crate::formula::omml::properties::parse_nary_properties;
use quick_xml::events::BytesStart;

/// Handler for n-ary operator elements
pub struct NaryHandler;

impl NaryHandler {
    pub fn handle_start<'arena>(
        elem: &BytesStart,
        context: &mut ElementContext<'arena>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let attrs: Vec<_> = elem.attributes().filter_map(|a| a.ok()).collect();

        // Parse n-ary properties
        context.properties = parse_nary_properties(&attrs);

        // Parse operator character using SIMD-accelerated parsing
        let chr_val = get_attribute_value(&attrs, "chr");
        if let Some(chr) = chr_val {
            context.properties.chr = Some(chr);
            context.operator = parse_large_operator(Some(&context.properties.chr.as_ref().unwrap()));
        }
    }

    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        // Get operator from properties (set by chr child element)
        let operator = context.properties.chr.as_ref()
            .and_then(|chr| parse_large_operator(Some(chr)))
            .or(context.operator);

        if let Some(operator) = operator {
            let lower_limit = context.lower_limit.take();
            let upper_limit = context.upper_limit.take();
            let integrand = context.integrand.take();

            let hide_lower = context.properties.nary_hide_sub.unwrap_or(false);
            let hide_upper = context.properties.nary_hide_sup.unwrap_or(false);

            let node = MathNode::LargeOp {
                operator,
                lower_limit,
                upper_limit,
                integrand,
                hide_lower,
                hide_upper,
            };

            if let Some(parent) = parent_context {
                parent.children.push(node);
            }
        }
    }
}
