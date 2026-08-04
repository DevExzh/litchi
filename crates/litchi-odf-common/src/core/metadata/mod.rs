//! ODF metadata owner layered by semantic model, bounded XML codec, and retained-source patching.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{
    AutoReloadMetadata, DocumentStatistics, HyperlinkBehaviourMetadata, Metadata, TemplateMetadata,
    UserDefinedMetadata, UserDefinedValueType,
};

pub use codec::{MetaXmlPatch, patch_meta_xml};
