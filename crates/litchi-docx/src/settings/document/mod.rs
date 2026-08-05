//! Document settings with layered model, XML codec, and OPC seams.

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub use model::{AttachedTemplate, DocumentSettings};

#[cfg(any(feature = "fonts", test))]
pub(crate) use codec::patch_font_embedding;
pub(crate) use codec::{patch_attached_template, patch_document_variables, patch_mail_merge};
pub(crate) use package::{
    ATTACHED_TEMPLATE_RELATIONSHIP, extract_document_variables, validate_attached_template_target,
};

#[cfg(test)]
pub(crate) use package::{
    STRICT_ATTACHED_TEMPLATE_RELATIONSHIP, is_attached_template_relationship,
};
