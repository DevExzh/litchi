//! Common OOXML functionality shared across formats.

pub mod mce;
pub mod properties;
pub(crate) mod xml;

pub use mce::{
    ExpandedName, MceCapabilities, MceError, MceLimits, MceOutput, MceReport,
    process_markup_compatibility,
};
pub use properties::DocumentProperties;
