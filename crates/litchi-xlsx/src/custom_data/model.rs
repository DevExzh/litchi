//! XLSX facade for the shared OOXML Custom Data values.
//!
//! Package graph state remains in `package.rs`; these names are re-exports so
//! existing XLSX callers keep their imports while XLSB can consume the same
//! common leaf without depending on this crate.

pub use litchi_ooxml_common::custom_data::{
    CustomData, CustomDataView, ExtensionList, Properties, RemovalDisposition, Store,
};
