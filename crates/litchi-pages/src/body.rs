//! The rooted body story as one rich-text value (record 0371).
//!
//! [`crate::Package::body_storage`] reads the body's complete native text
//! (section breaks, object-replacement and footnote markers included) with
//! its style and attribute runs. The value is the shared iWork rich-text
//! storage: it holds no native object identifier, style-table identifier or
//! protobuf value.

pub use litchi_iwa_text::storage::{Run, Storage};
