//! Error conversion implementations.
//!
//! This module contains `From` trait implementations for error types whose
//! source crates are reachable from `litchi-core` (quick-xml, soapberry-zip).
//! Conversions for per-format error types (`crate::ole::*`, `crate::ooxml::*`)
//! live in the umbrella crate at `src/error_ext.rs` because their source
//! types are not visible to `litchi-core`.

use super::types::Error;

impl From<quick_xml::Error> for Error {
    fn from(err: quick_xml::Error) -> Self {
        Error::XmlError(err.to_string())
    }
}

impl From<soapberry_zip::Error> for Error {
    fn from(err: soapberry_zip::Error) -> Self {
        Error::ZipError(err.to_string())
    }
}
