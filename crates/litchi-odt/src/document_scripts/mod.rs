//! Inert ODF document script declarations.
 //!
 //! This module intentionally exposes script payloads as metadata only. It never
 //! loads linked resources and never executes embedded script content.

mod codec;
mod model;
pub(crate) mod package;
#[cfg(test)]
mod tests;

use litchi_core::{Error, Result};

pub(super) const OFFICE_NAMESPACE: &[u8] =
    b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const SCRIPT_NAMESPACE: &[u8] =
    b"urn:oasis:names:tc:opendocument:xmlns:script:1.0";
pub(super) const PRESENTATION_NAMESPACE: &[u8] =
    b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
pub(super) const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
pub(super) const MAX_DOCUMENT_XML_BYTES: usize = 32 * 1024 * 1024;
pub(super) const MAX_SCRIPT_COUNT: usize = 1_024;
pub(super) const MAX_LISTENER_COUNT: usize = 4_096;
pub(super) const MAX_TEXT_BYTES: usize = 8 * 1024 * 1024;
pub(super) const MAX_VALUE_BYTES: usize = 64 * 1024;
pub(super) const MAX_XML_DEPTH: usize = 128;

pub(super) fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

pub use codec::parse_scripts;
pub use model::{EmbeddedScript, EventListener, ScriptBinding, ScriptEventListener, Scripts};
