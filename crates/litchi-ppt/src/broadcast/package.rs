//! Document-level PowerPoint 9 broadcast collection.

use super::codec::BROADCAST_CONTAINER_RECORD_TYPE;
use super::model::{Broadcast, Broadcasts};
use crate::package::Result;
use crate::records::Record;

impl Broadcasts {
    pub(crate) fn parse_document(document: &Record) -> Result<Self> {
        let mut broadcasts = Vec::new();
        for record in document.versioned_binary_tag_records(9)? {
            if record.record_type_raw == BROADCAST_CONTAINER_RECORD_TYPE {
                broadcasts.push(Broadcast::parse(&record)?);
            }
        }
        Ok(Self { broadcasts })
    }
}
