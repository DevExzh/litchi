// Limit element handler

use crate::omml::elements::{ElementContext, ElementType};

/// Handler for limit elements
pub struct LimitHandler;

impl LimitHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: limit elements are owned Vec from context
    ) {
        if let Some(parent) = parent_context {
            match parent.element_type {
                // The m:lim child of m:limLow / m:limUpp carries the script
                ElementType::LimLow => {
                    parent.lower_limit = Some(context.children.clone());
                },
                ElementType::LimUpp => {
                    parent.upper_limit = Some(context.children.clone());
                },
                _ => {
                    crate::omml::utils::extend_vec_efficient(
                        &mut parent.children,
                        context.children.clone(),
                    );
                },
            }
        }
    }
}
