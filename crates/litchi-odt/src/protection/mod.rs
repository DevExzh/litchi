//! Contextual ODT document interaction and protection metadata.
//!
//! The owner is intentionally narrower than an enforcement engine. It maps
//! the producer-visible protection settings used by ODF text documents onto a
//! small typed policy, while retaining every unmodeled settings node and the
//! original XML bytes during edits.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{Key, Policy};
pub use transaction::{
    Commit, DurablePatch, Field, MergePlan, Patch, Resolution, Transaction, Transfer,
};

pub(crate) use codec::{parse_flat, parse_package};
pub(crate) use package::rewrite_owned_package;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const CONFIG_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:config:1.0";
const OFFICE_NAMESPACE_TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const CONFIG_NAMESPACE_TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
const CONFIGURATION_SET: &str = "ooo:configuration-settings";
const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_VALUE_BYTES: usize = 4 * 1024 * 1024;
const MAX_KEY_BYTES: usize = 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_ITEMS: usize = 65_536;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Kind {
    Flat,
    Package,
}

pub(crate) fn invalid<T>(message: impl Into<String>) -> litchi_core::Result<T> {
    Err(litchi_core::Error::InvalidFormat(message.into()))
}

pub(crate) fn xml_error(error: impl std::fmt::Display) -> litchi_core::Error {
    litchi_core::Error::InvalidFormat(format!("invalid ODT protection settings XML: {error}"))
}
