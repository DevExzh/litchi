//! Inert mail-merge package payloads and recipient-part extraction.

use super::RECIPIENT_CONTENT_TYPE;
use super::model::{Recipients, invalid};
use crate::Result;
use litchi_opc::part::Part;

/// Opaque mail-merge source to relate from `settings.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Source {
    /// Bytes are stored as an inert package part and never opened or interpreted.
    Internal {
        bytes: Vec<u8>,
        content_type: String,
        extension: String,
    },
    /// URI is stored as an external relationship and never fetched.
    External(String),
}

/// Owned, inert relationship target returned by package lookup.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Target {
    Internal {
        part_name: litchi_opc::PackURI,
        bytes: Vec<u8>,
        content_type: String,
    },
    External(String),
}

impl Recipients {
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn extract_from_part(part: &dyn Part) -> Result<Self> {
        if part.content_type() != RECIPIENT_CONTENT_TYPE {
            return Err(invalid(format!(
                "invalid mail-merge recipient-data content type '{}'",
                part.content_type()
            )));
        }
        if part.rels().iter().next().is_some() {
            return Err(invalid(
                "mail-merge recipient-data part cannot have relationships",
            ));
        }
        let xml = litchi_ooxml_common::mce::process_part(part)?;
        Self::parse_xml(xml.as_ref())
    }
}
