// Equation array element handler

use crate::formula::ast::*;
use crate::formula::omml::elements::ElementContext;
use crate::formula::omml::attributes::{get_attribute_value, get_attribute_value_int, get_attribute_value_float};
use crate::formula::omml::properties::parse_eq_arr_properties;
use quick_xml::events::BytesStart;

/// Handler for equation array elements
pub struct EqArrHandler;

impl EqArrHandler {
    pub fn handle_start<'arena>(
        elem: &BytesStart,
        context: &mut ElementContext<'arena>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let attrs: Vec<_> = elem.attributes().filter_map(|a| a.ok()).collect();

        // Parse equation array properties
        context.properties = parse_eq_arr_properties(&attrs);

        // Parse additional attributes using SIMD-accelerated parsing
        if let Some(max_dist) = get_attribute_value_int(&attrs, "maxDist") {
            // Store as string for now, will be parsed in handle_end
            context.properties.eq_arr_max_distance = Some(max_dist.to_string());
        }

        if let Some(obj_dist) = get_attribute_value_int(&attrs, "objDist") {
            context.properties.eq_arr_object_distance = Some(obj_dist.to_string());
        }

        if let Some(r_sp) = get_attribute_value_float(&attrs, "rSp") {
            context.properties.eq_arr_row_spacing = Some(r_sp.to_string());
        }

        let base_jc_val = get_attribute_value(&attrs, "baseJc");
        if let Some(base_jc) = base_jc_val {
            context.properties.eq_arr_base_alignment = Some(base_jc);
        }

        let r_sp_rule_val = get_attribute_value(&attrs, "rSpRule");
        if let Some(r_sp_rule) = r_sp_rule_val {
            context.properties.eq_arr_row_spacing_rule = Some(r_sp_rule);
        }
    }

    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let rows = std::mem::take(&mut context.eq_array_rows);

        // Create equation array properties from context
        let properties = if context.properties.eq_arr_base_alignment.is_some()
            || context.properties.eq_arr_max_distance.is_some()
            || context.properties.eq_arr_object_distance.is_some()
            || context.properties.eq_arr_row_spacing.is_some()
            || context.properties.eq_arr_row_spacing_rule.is_some() {
            Some(EqArrayProperties {
                base_alignment: context.properties.eq_arr_base_alignment
                    .as_ref()
                    .and_then(|s| match s.as_str() {
                        "top" => Some(Alignment::Top),
                        "center" | "cen" => Some(Alignment::Center),
                        "bottom" | "bot" => Some(Alignment::Bottom),
                        _ => None,
                    }),
                max_distance: context.properties.eq_arr_max_distance
                    .as_ref()
                    .and_then(|s| s.parse().ok()),
                object_distance: context.properties.eq_arr_object_distance
                    .as_ref()
                    .and_then(|s| s.parse().ok()),
                row_spacing: context.properties.eq_arr_row_spacing
                    .as_ref()
                    .and_then(|s| s.parse().ok()),
                row_spacing_rule: context.properties.eq_arr_row_spacing_rule.clone(),
            })
        } else {
            None
        };

        // Create equation array node
        let node = MathNode::EqArray { rows, properties };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}
