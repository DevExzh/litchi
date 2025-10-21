// Fraction element handler

use crate::formula::ast::*;
use crate::formula::omml::elements::ElementContext;
use crate::formula::omml::properties::parse_fraction_properties;
use quick_xml::events::BytesStart;

/// Handler for fraction elements
pub struct FractionHandler;

impl FractionHandler {
    pub fn handle_start<'arena>(
        elem: &BytesStart,
        context: &mut ElementContext<'arena>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let attrs: Vec<_> = elem.attributes().filter_map(|a| a.ok()).collect();

        // Parse fraction properties
        context.properties = parse_fraction_properties(&attrs);
    }

    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        let numerator = context.numerator.take().unwrap_or_default();
        let denominator = context.denominator.take().unwrap_or_default();

        let line_thickness = context
            .properties
            .fraction_line_thickness
            .as_ref()
            .and_then(|s| s.parse().ok());

        let frac_type = context
            .properties
            .fraction_type
            .as_ref()
            .and_then(|s| match s.as_str() {
                "bar" => Some(FractionType::Bar),
                "noBar" => Some(FractionType::NoBar),
                "skw" | "skewed" => Some(FractionType::Skewed),
                _ => None,
            });

        let node = MathNode::Frac {
            numerator,
            denominator,
            line_thickness,
            frac_type,
        };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}
