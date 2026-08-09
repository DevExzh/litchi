//! ODF metadata owner layered by semantic model, bounded XML codec, and retained-source patching.

mod codec;
mod model;

#[cfg(test)]
mod tests;

#[allow(
    clippy::module_name_repetitions,
    reason = "These names are the public ODF metadata vocabulary and mirror the format."
)]
pub use model::{
    AutoReloadMetadata, DocumentStatistics, HyperlinkBehaviourMetadata, Metadata, TemplateMetadata,
    UserDefinedMetadata, UserDefinedValueType,
};

#[allow(
    clippy::module_name_repetitions,
    reason = "`MetaXmlPatch` identifies the public meta.xml patch type."
)]
pub use codec::{MetaXmlPatch, patch_meta_xml};
