// Matrix cell element handler

use crate::formula::omml::elements::ElementContext;

/// Handler for matrix cell elements
#[allow(dead_code)] // Handler implementation, used by matrix row handler
pub struct MatrixCellHandler;

#[allow(dead_code)] // Handler implementation
impl MatrixCellHandler {
    pub fn handle_end<'arena>(
        context: &mut ElementContext<'arena>,
        parent_context: Option<&mut ElementContext<'arena>>,
        _arena: &'arena bumpalo::Bump,
    ) {
        if let Some(parent) = parent_context {
            // Matrix cells are processed by the matrix row handler
            // Here we just pass children up to the matrix row
            crate::formula::omml::utils::extend_vec_efficient(&mut parent.children, context.children.clone());
        }
    }
}
