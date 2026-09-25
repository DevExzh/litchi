//! Shared structures independent of DOC, PPT, and XLS semantic models.

#![forbid(unsafe_code)]
// quick-xml's checked attribute iteration is quadratic on hostile tags; read
// attributes through `BytesStartExt` (record 0770, workspace `clippy.toml`).
#![cfg_attr(not(test), deny(clippy::disallowed_methods))]

pub mod custom_xml;
pub mod dataspaces;
pub mod object;
pub mod ole1;
pub mod ole_streams;
pub mod property_set;
pub mod protection;
pub mod smart_tags;
pub mod source_backed_overlay;
pub mod toolbar;
pub mod vba_signature;
pub mod xml_attributes;
