//! Host-neutral Custom XML Data Storage package grammar.
//!
//! [`Snapshot::load`] reads the graph; [`Edit::new`] stages a change over a
//! shared package whose [`Edit::commit`] publishes a new [`litchi_opc::OpcPackage`]
//! in a [`Commit`] with its exact [`Patch`] and [`Diagnostics`];
//! [`Patch::apply`] replays a patch the same way. Budgets are typed by
//! [`LimitKind`].
//!
//! Payload XML is validated and retained as inert bytes. This module never
//! retrieves schemas, validates against a schema, runs transforms, resolves
//! external entities, or interprets application-specific payloads.

mod codec;
mod edit;
mod errors;
mod limit;
mod model;
mod package;
mod patch;
mod snapshot;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{
    read_props, valid_guid, validate_content_type, validate_payload, validate_props, write_props,
};
pub use edit::Edit;
pub use limit::LimitKind;
pub use model::{
    Conformance, Item, MAX_DEPTH, MAX_ELEMENTS, MAX_ITEMS, MAX_PART_BYTES, MAX_PROPS_BYTES,
    MAX_SCHEMA_REFS, MAX_STRING_BYTES, NewItem, NewProps, PROPS_CONTENT_TYPE, Props, Relationship,
    STRICT_NAMESPACE, STRICT_PROPS_RELATIONSHIP, STRICT_RELATIONSHIP, TRANSITIONAL_NAMESPACE,
    TRANSITIONAL_PROPS_RELATIONSHIP, TRANSITIONAL_RELATIONSHIP,
};
pub use package::discover;
pub use patch::{Commit, Diagnostics, Patch};
pub use snapshot::Snapshot;

/// The error of the Custom XML family: the crate's error.
pub use crate::Error;
