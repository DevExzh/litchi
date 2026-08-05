//! Worksheet chart integration for SpreadsheetML.
//!
//! The facade keeps host-facing names contextual: [`Chart`] is an XLSX chart
//! placement and resource bundle, while the chart payload itself remains the
//! shared [`litchi_drawingml::chart`] model. Relationship resources, worksheet
//! anchors, and XML codecs have dedicated submodule paths so package code does
//! not need to depend on the model implementation module.

#![allow(dead_code)]

pub mod anchor;
pub mod codec;
pub mod model;
pub mod relationship;

pub use anchor::Anchor;
pub use codec::{generate_chart_xml, parse_chart_from_xml};
pub use model::{Chart, Series};
pub use relationship::{
    ExternalDataPart, ExternalDataTarget, Relationship, RelationshipTarget, UserShapesPart,
};

#[cfg(test)]
mod tests;
