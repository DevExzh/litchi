//! Shared Feature12 paths, including bounded opaque table features.

use crate::Result;
use crate::list_object::codec::binary::{
    PendingFeature, parse_range, parse_string, u16_at, u32_at, validate_frt,
};
use crate::list_object::model::{
    ListObject, ListObjectFeatureVersion, ListObjectId, OpaqueListObjectFeature,
};
use crate::list_object::{FEATURE12_RECORD_TYPE, ISF_LIST, invalid};

impl ListObject {
    pub(in crate::list_object) fn parse_opaque_feature12(pending: PendingFeature) -> Result<Self> {
        let data = &pending.combined;
        if data.len() < 108 {
            return Err(invalid(FEATURE12_RECORD_TYPE, "truncated opaque Feature12"));
        }
        validate_frt(data, FEATURE12_RECORD_TYPE, true)?;
        if u16_at(data, 12, FEATURE12_RECORD_TYPE, "isf")? != ISF_LIST
            || data[14] != 0
            || u32_at(data, 55, FEATURE12_RECORD_TYPE, "cbFSData")? != 64
        {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "invalid opaque Feature12 fixed fields",
            ));
        }
        let range = parse_range(data, 4, FEATURE12_RECORD_TYPE)?;
        let id = ListObjectId::try_new(u32_at(data, 39, FEATURE12_RECORD_TYPE, "idList")?)?;
        let header = u32_at(data, 43, FEATURE12_RECORD_TYPE, "crwHeader")?;
        let totals = u32_at(data, 47, FEATURE12_RECORD_TYPE, "crwTotals")?;
        if header > 1 || totals > 1 {
            return Err(invalid(
                FEATURE12_RECORD_TYPE,
                "invalid opaque Feature12 row flags",
            ));
        }
        let flags = u32_at(data, 63, FEATURE12_RECORD_TYPE, "flags")?;
        let table_flags = crate::list_object::TableFlags::from_raw(flags);
        let (name, _) = parse_string(data, 99, FEATURE12_RECORD_TYPE, "rgbName")?;
        let opaque_feature = OpaqueListObjectFeature {
            record_type: pending.record_type,
            base_payload: pending.base,
            continuation_payloads: pending.continuations,
        };
        Ok(Self {
            id,
            name,
            range,
            columns: Vec::new(),
            style: None,
            has_header: header != 0,
            has_totals: totals != 0,
            autofilter: table_flags.auto_filter(),
            table_flags,
            comment: String::new(),
            feature_version: ListObjectFeatureVersion::Feature12,
            opaque_feature: Some(opaque_feature),
            opaque_future_records: Vec::new(),
            autofilter12_criteria: None,
            external_metadata: None,
            source_metadata: None,
        })
    }
}
