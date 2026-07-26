// Pre-script element handler

use crate::ast::MathNode;
use crate::omml::elements::ElementContext;

/// Handler for pre-script elements (m:sPre, ECMA-376 §22.1.2.99)
///
/// Builds `PreSub` / `PreSup` / `PreSubSup` nodes from the `m:sub`, `m:sup`
/// and `m:e` children.
pub struct PreScriptHandler;

impl PreScriptHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump, // Unused: script content is owned Vec from context
    ) {
        let Some(parent) = parent_context else {
            return;
        };

        let base = context.base.take();
        let subscript = context.subscript.take().filter(|s| !s.is_empty());
        let superscript = context.superscript.take().filter(|s| !s.is_empty());

        match (subscript, superscript) {
            (Some(pre_subscript), Some(pre_superscript)) => {
                parent.children.push(MathNode::PreSubSup {
                    base: base.unwrap_or_default(),
                    pre_subscript,
                    pre_superscript,
                });
            },
            (Some(pre_subscript), None) => {
                parent.children.push(MathNode::PreSub {
                    base: base.unwrap_or_default(),
                    pre_subscript,
                });
            },
            (None, Some(pre_superscript)) => {
                parent.children.push(MathNode::PreSup {
                    base: base.unwrap_or_default(),
                    pre_superscript,
                });
            },
            (None, None) => {
                // No scripts present: pass any content up unchanged
                if let Some(base) = base {
                    crate::omml::utils::extend_vec_efficient(&mut parent.children, base);
                }
                crate::omml::utils::extend_vec_efficient(
                    &mut parent.children,
                    context.children.clone(),
                );
            },
        }
    }
}
