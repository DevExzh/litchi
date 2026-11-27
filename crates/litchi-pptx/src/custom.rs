//! Inert custom document properties for `PresentationML` packages.
//!
//! Values and package-graph semantics are shared by every OOXML format. This
//! contextual module exposes the common typed vocabulary without duplicating
//! its XML codec or OPC ownership.

pub use litchi_ooxml_common::custom::{Host, Props, Value};
