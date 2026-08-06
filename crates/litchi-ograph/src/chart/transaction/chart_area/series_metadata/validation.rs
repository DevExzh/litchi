//! `[MS-OGRAPH]` semantic guards for the fixed-width `Series` record.

use super::Metadata;
use crate::{Error, Result};

const MAX_COUNT: u16 = 0x0F9F;

pub(super) fn ensure(metadata: Metadata) -> Result<()> {
    for (field, value) in [
        ("category count", metadata.category_count().get()),
        ("value count", metadata.value_count().get()),
        ("bubble count", metadata.bubble_count().get()),
    ] {
        if value > MAX_COUNT {
            return Err(Error::InvalidModel {
                field,
                reason: "Series point count exceeds the MS-OGRAPH limit 0x0F9F",
            });
        }
    }
    Ok(())
}
