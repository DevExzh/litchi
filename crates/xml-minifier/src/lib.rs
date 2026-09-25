//! Development-time XML compactness enforcement and template macros.
//!
//! The auditor never rewrites input. It parses XML, reports the first proven
//! compactness defect, and enforces finite resource budgets. The macros remain
//! available at their historical paths for producer-template regeneration.
//!
//! The compactness contract binds bytes this repository authors.
//! [`audit::verify_source`] is the policy for bytes it did not author: it
//! keeps every structural, encoding, DOCTYPE and budget check and asserts no
//! compactness, so a producer's whitespace, XML declaration, line endings and
//! attribute spelling are accepted as written. [`audit::verify_source_replacement`]
//! gives the verdict of that audit on a replaced payload's original bytes and
//! then its replacement, without scanning again the replacement bytes that
//! repeat the original's outside the one element an edit replaced.

#![forbid(unsafe_code)]
// quick-xml's checked attribute iteration is quadratic on hostile tags; read
// attributes through `BytesStartExt` (record 0770, workspace `clippy.toml`).
#![cfg_attr(not(test), deny(clippy::disallowed_methods))]

pub mod audit;
mod xml_attributes;

pub use xml_minifier_macros::{minified_xml, minified_xml_format, minified_xml_str};
