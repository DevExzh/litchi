//! Common OOXML functionality shared across formats.

pub mod properties;
pub mod mce;
pub(crate) mod xml;

pub use properties::DocumentProperties;
pub use mce::{ExpandedName,MceCapabilities,MceError,MceLimits,MceOutput,MceReport,process_markup_compatibility};
