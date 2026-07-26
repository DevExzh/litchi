// Handler for single-value property elements
//
// OMML expresses most properties as child elements that carry a single
// `m:val` attribute, e.g. `<m:type m:val="noBar"/>` inside `m:fPr` or
// `<m:subHide m:val="1"/>` inside `m:naryPr` (ECMA-376 Part 1, §22.1).
// This handler stores the value into the parent property context based on
// the element name.

use crate::omml::elements::{ElementContext, ElementType};
use crate::omml::lookup::parse_bool_value;

/// Handler for property elements that carry a single value
pub struct ValPropertyHandler;

impl ValPropertyHandler {
    pub fn handle_end<'arena>(
        elem: &[u8],
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: values stored as owned strings in parent properties
    ) {
        let Some(parent) = parent_context else {
            return;
        };

        // Value comes from the m:val attribute (captured during element start)
        // or, as a fallback, from the element text content.
        let value = context.character_data.take().or_else(|| {
            if context.text.is_empty() {
                None
            } else {
                Some(context.text.as_str().trim().to_string())
            }
        });

        let Some(value) = value else {
            return;
        };

        let name = std::str::from_utf8(elem).unwrap_or("");
        match name {
            "type" | "m:type" => {
                parent.properties.fraction_type = Some(value);
            },
            "subHide" | "m:subHide" => {
                parent.properties.nary_hide_sub = parse_bool_value(&value);
            },
            "supHide" | "m:supHide" => {
                parent.properties.nary_hide_sup = parse_bool_value(&value);
            },
            "baseJc" | "m:baseJc" => {
                if parent.element_type == ElementType::EqArrPr {
                    parent.properties.eq_arr_base_alignment = Some(value);
                } else {
                    parent.properties.matrix_alignment = Some(value);
                }
            },
            "maxDist" | "m:maxDist" => {
                parent.properties.eq_arr_max_distance = Some(value);
            },
            "objDist" | "m:objDist" => {
                parent.properties.eq_arr_object_distance = Some(value);
            },
            "rSp" | "m:rSp" => {
                if parent.element_type == ElementType::EqArrPr {
                    parent.properties.eq_arr_row_spacing = Some(value);
                } else {
                    parent.properties.matrix_row_spacing = Some(value);
                }
            },
            "rSpRule" | "m:rSpRule" => {
                parent.properties.eq_arr_row_spacing_rule = Some(value);
            },
            _ => {},
        }
    }
}
