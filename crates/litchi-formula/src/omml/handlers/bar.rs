// Bar element handler

use crate::ast::*;
use crate::omml::attributes::parse_position_type;
use crate::omml::elements::ElementContext;

/// Handler for bar elements
pub struct BarHandler;

impl BarHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: simple wrapper, children are owned Vec
    ) {
        let base = if context.children.is_empty() {
            Vec::new()
        } else {
            context.children.clone()
        };

        // Use dedicated position parsing function. The position may come from
        // an `m:pos` child element in `m:barPr` (accent_position) or from an
        // alignment attribute.
        let position = context
            .properties
            .accent_position
            .as_ref()
            .or(context.properties.alignment.as_ref())
            .and_then(|s| parse_position_type(Some(s)));

        let node = MathNode::Bar {
            base: Box::new(base),
            position,
        };

        if let Some(parent) = parent_context {
            parent.children.push(node);
        }
    }
}
