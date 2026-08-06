//! ODF table-cell protection styles with a stable, prefix-free facade.

mod codec;
mod model;
mod semantic;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{ConditionalStyle, Protection, Rule, TableStyle};

#[allow(unused_imports)]
pub(crate) use codec::{
    extract_automatic_styles, extract_font_face_decls, rewrite_conditional_styles,
    rewrite_managed_cell_styles,
};
#[allow(unused_imports)]
pub(crate) use model::{CellStyleRegistry, PreservedXmlFragment};
#[allow(unused_imports)]
pub(crate) use semantic::{common_table_cell_style_names, validate_protection_style_document};
#[allow(unused_imports)]
pub(crate) use validation::{
    validate_conditional_style_collection, validate_protection_style_collection,
    validate_style_name,
};

pub(super) const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
