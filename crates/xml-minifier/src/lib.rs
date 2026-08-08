//! Development-time XML compactness enforcement and template macros.
//!
//! The auditor never rewrites input. It parses XML, reports the first proven
//! compactness defect, and enforces finite resource budgets. The macros remain
//! available at their historical paths for producer-template regeneration.

#![forbid(unsafe_code)]

pub mod audit;

pub use xml_minifier_macros::{minified_xml, minified_xml_format, minified_xml_str};
