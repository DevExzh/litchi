//! Immutable spreadsheet-domain vocabulary.
#![allow(
    dead_code,
    reason = "model codecs are retained for the package parser migration"
)]

pub mod conditional_format;
pub mod consolidation;
pub mod data_pilot;
pub mod data_validation;
pub mod database_range;
pub mod detective;
pub mod hyperlink;
pub mod label_range;
pub mod names;
pub mod protection;
pub mod source;
pub mod sparkline;
pub mod structure;
pub mod style_protection;
pub mod table_template;
pub mod tracked_changes;

pub use litchi_odf_common::{annotation, calculation, rdf};
