//! Shared Feature11/Feature12 dispatch and record ownership.

use crate::Result;
use crate::list_object::CONTINUE_FRT11_RECORD_TYPE;
use crate::list_object::codec::binary::record;
use crate::list_object::model::ListObject;

impl ListObject {
    pub(crate) fn to_feature_record_bytes(&self) -> Result<Vec<Vec<u8>>> {
        self.validate()?;
        if let Some(opaque) = &self.opaque_feature {
            let mut records = vec![record(opaque.record_type, opaque.base_payload.clone())?];
            for payload in &opaque.continuation_payloads {
                records.push(record(CONTINUE_FRT11_RECORD_TYPE, payload.clone())?);
            }
            return Ok(records);
        }
        if let Some(metadata) = &self.external_metadata {
            return self.to_external_feature_record_bytes(metadata);
        }
        if let Some(metadata) = &self.source_metadata {
            return self.to_source_feature_record_bytes(metadata);
        }
        self.to_ordinary_feature_record_bytes()
    }
}
