//! Host-neutral Custom XML Data Storage package grammar.
//!
//! Payload XML is validated and retained as inert bytes. This module never
//! retrieves schemas, validates against a schema, runs transforms, resolves
//! external entities, or interprets application-specific payloads.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::{
    read_props, valid_guid, validate_content_type, validate_payload, validate_props, write_props,
};
pub use model::{
    Conformance, Item, MAX_DEPTH, MAX_ELEMENTS, MAX_ITEMS, MAX_PART_BYTES, MAX_PROPS_BYTES,
    MAX_SCHEMA_REFS, MAX_STRING_BYTES, NewItem, NewProps, PROPS_CONTENT_TYPE, Props,
    STRICT_NAMESPACE, STRICT_PROPS_RELATIONSHIP, STRICT_RELATIONSHIP, TRANSITIONAL_NAMESPACE,
    TRANSITIONAL_PROPS_RELATIONSHIP, TRANSITIONAL_RELATIONSHIP,
};
pub use package::{add, discover};
