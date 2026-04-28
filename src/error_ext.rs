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

/// Convert an OOXML OPC error into the unified `Error` type.
///
/// Inlined here because litchi-core's previous `Error::from_opc_error` helper
/// was `pub(crate)` and is no longer reachable across the crate boundary.
#[cfg(feature = "ooxml")]
fn from_opc_error(err: crate::ooxml::opc::error::OpcError) -> Error {
    match err {
        crate::ooxml::opc::error::OpcError::IoError(e) => Error::Io(e),
        crate::ooxml::opc::error::OpcError::ZipError(e) => Error::ZipError(e.to_string()),
        crate::ooxml::opc::error::OpcError::XmlError(s) => Error::XmlError(s),
        crate::ooxml::opc::error::OpcError::PartNotFound(s) => Error::ComponentNotFound(s),
        _ => Error::Other(err.to_string()),
    }
}

#[cfg(feature = "ooxml")]
impl From<crate::ooxml::opc::error::OpcError> for Error {
    fn from(err: crate::ooxml::opc::error::OpcError) -> Self {
        from_opc_error(err)
    }
}

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
            crate::ooxml::error::OoxmlError::Opc(e) => from_opc_error(e),
            crate::ooxml::error::OoxmlError::IoError(e) => Error::Io(e),
            crate::ooxml::error::OoxmlError::InvalidUri(s) => Error::Other(s),
            crate::ooxml::error::OoxmlError::Other(s) => Error::Other(s),
        }
    }
}
