//! Umbrella-side `From` implementations that convert per-format error types
//! (`crate::ole::*`, `crate::ooxml::*`) into the unified `litchi_core::Error`.
//!
//! These impls cannot live inside `litchi-core` because they reference
//! umbrella-local types (`crate::ole`, `crate::ooxml`). The orphan rules
//! permit this here because the *source* type of each `From` is local to the
//! umbrella crate.

#[cfg(any(feature = "ole", feature = "ooxml"))]
use litchi_core::Error;

// ---------------------------------------------------------------------------
// OLE error conversions
// ---------------------------------------------------------------------------
//
// Note: `From<crate::ole::OleError> for Error` is provided by the `litchi-cfb`
// crate itself (since `OleError` was carved out into that crate, the orphan
// rule forbids us from implementing the conversion at the umbrella). The doc/
// ppt package-level conversions below remain here because their source types
// (`DocError`, `PptError`) are still local to the umbrella.

#[cfg(feature = "ole")]
impl From<crate::ole::doc::package::DocError> for Error {
    fn from(err: crate::ole::doc::package::DocError) -> Self {
        match err {
            crate::ole::doc::package::DocError::Io(e) => Error::Io(e),
            crate::ole::doc::package::DocError::Ole(ole_err) => Error::from(ole_err),
            crate::ole::doc::package::DocError::InvalidFormat(s) => Error::InvalidFormat(s),
            crate::ole::doc::package::DocError::StreamNotFound(s) => Error::ComponentNotFound(s),
            crate::ole::doc::package::DocError::Corrupted(s) => Error::CorruptedFile(s),
        }
    }
}

#[cfg(feature = "ole")]
impl From<crate::ole::ppt::package::PptError> for Error {
    fn from(err: crate::ole::ppt::package::PptError) -> Self {
        match err {
            crate::ole::ppt::package::PptError::Io(e) => Error::Io(e),
            crate::ole::ppt::package::PptError::Ole(ole_err) => Error::from(ole_err),
            crate::ole::ppt::package::PptError::InvalidFormat(s) => Error::InvalidFormat(s),
            crate::ole::ppt::package::PptError::StreamNotFound(s) => Error::ComponentNotFound(s),
            crate::ole::ppt::package::PptError::Corrupted(s) => Error::CorruptedFile(s),
        }
    }
}

// ---------------------------------------------------------------------------
// OOXML error conversions
// ---------------------------------------------------------------------------

// `impl From<OpcError> for litchi_core::Error` (and the private `from_opc_error`
// helper it delegated to) moved to `crates/litchi-opc/src/error.rs` to satisfy
// the orphan rule after the litchi-opc carve-out (P3b). Keep this comment for
// grep traceability.

#[cfg(feature = "ooxml")]
impl From<crate::ooxml::error::OoxmlError> for Error {
    fn from(err: crate::ooxml::error::OoxmlError) -> Self {
        match err {
            crate::ooxml::error::OoxmlError::Io(e) => Error::Io(e),
            crate::ooxml::error::OoxmlError::Xml(s) => Error::XmlError(s),
            crate::ooxml::error::OoxmlError::PartNotFound(s) => Error::ComponentNotFound(s),
            crate::ooxml::error::OoxmlError::InvalidContentType { expected, got } => {
                Error::InvalidContentType { expected, got }
            },
            crate::ooxml::error::OoxmlError::InvalidRelationship(s) => Error::Other(s),
            crate::ooxml::error::OoxmlError::InvalidFormat(s) => Error::InvalidFormat(s),
            crate::ooxml::error::OoxmlError::Opc(e) => Error::from(e),
            crate::ooxml::error::OoxmlError::IoError(e) => Error::Io(e),
            crate::ooxml::error::OoxmlError::InvalidUri(s) => Error::Other(s),
            crate::ooxml::error::OoxmlError::Other(s) => Error::Other(s),
        }
    }
}
