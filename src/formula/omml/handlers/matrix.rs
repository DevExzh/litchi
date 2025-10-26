// Matrix element handlers

use crate::formula::ast::*;
use crate::formula::omml::attributes::{get_attribute_value, parse_matrix_fence};
use crate::formula::omml::elements::{ElementContext, ElementType};
use crate::formula::omml::properties::parse_matrix_properties;
use quick_xml::events::BytesStart;

/// Handler for matrix elements
pub struct MatrixHandler;

impl MatrixHandler {
    pub fn handle_start<'arena>(
        elem: &BytesStart,
        context: &mut ElementContext<'arena>,
        _arena: &'arena bumpalo::Bump, // Unused: matrix elements are owned Vec, no string allocation
    ) {
        let attrs: Vec<_> = elem.attributes().filter_map(|a| a.ok()).collect();

        // Parse matrix column spacing (mcs) attribute using SIMD-accelerated parsing
        let fence_val = get_attribute_value(&attrs, "mcs");
        context.matrix_fence = parse_matrix_fence(fence_val.as_deref());

        // Parse matrix properties
        context.properties = parse_matrix_properties(&attrs);
    }

    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: matrix elements are owned Vec, no string allocation
    ) {
        let fence_type = context.matrix_fence.unwrap_or(MatrixFence::None);
        let rows = std::mem::take(&mut context.matrix_rows);

        // Create matrix properties from context
        let properties =
            if context.properties.matrix_alignment.is_some()
                || context.properties.matrix_row_spacing.is_some()
                || context.properties.matrix_column_spacing.is_some()
            {
                Some(MatrixProperties {
                    base_alignment: context.properties.matrix_alignment.as_ref().and_then(|s| {
                        match s.as_str() {
                            "top" => Some(Alignment::Top),
                            "center" | "cen" => Some(Alignment::Center),
                            "bottom" | "bot" => Some(Alignment::Bottom),
                            "baseline" | "base" => Some(Alignment::Baseline),
                            _ => None,
                        }
                    }),
                    column_gap: context
                        .properties
                        .matrix_column_spacing
                        .as_ref()
                        .and_then(|s| s.parse().ok()),
                    row_spacing: context
                        .properties
                        .matrix_row_spacing
                        .as_ref()
                        .and_then(|s| s.parse().ok()),
                    column_spacing: None, // Would need more complex parsing
                })
            } else {
                None
            };

        let node = MathNode::Matrix {
            rows,
            fence_type,
            properties,
        };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}

/// Handler for matrix row elements
pub struct MatrixRowHandler;

impl MatrixRowHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: matrix elements are owned Vec, no string allocation
    ) {
        if let Some(parent) = parent_context
            && parent.element_type == ElementType::Matrix
        {
            // Matrix row - collect cells from children
            // Each child represents a cell (mtd element)
            let mut row = Vec::new();
            for child in &context.children {
                // Each child is a cell containing mathematical content
                row.push(vec![child.clone()]);
            }
            parent.matrix_rows.push(row);
        }
    }
}
