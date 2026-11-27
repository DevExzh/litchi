//! Bounded validation shared by smart-tag snapshots and transactions.

use super::model::{Error, Limits, PropertyBag, PropertyBagStore};

pub(super) fn encode(
    store: &PropertyBagStore,
    bags: &[PropertyBag],
    limits: Limits,
) -> Result<Vec<u8>, Error> {
    if store.types.len() > limits.max_types {
        return Err(Error::new(
            "smart-tag type count exceeds the configured limit",
        ));
    }
    if store.strings.len() > limits.max_strings {
        return Err(Error::new(
            "smart-tag string count exceeds the configured limit",
        ));
    }
    if bags.len() > limits.max_bags {
        return Err(Error::new(
            "smart-tag bag count exceeds the configured limit",
        ));
    }

    let mut property_count = 0usize;
    for bag in bags {
        property_count = property_count
            .checked_add(bag.properties.len())
            .ok_or_else(|| Error::new("smart-tag property count overflows"))?;
    }
    if property_count > limits.max_properties {
        return Err(Error::new(
            "smart-tag property count exceeds the configured limit",
        ));
    }

    let bytes = store.to_bytes_with_bags(bags)?;
    if bytes.len() > limits.max_bytes {
        return Err(Error::new(
            "smart-tag serialized payload exceeds the configured limit",
        ));
    }
    Ok(bytes)
}
