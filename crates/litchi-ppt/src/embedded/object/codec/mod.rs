//! Record codecs for the typed OLE object models.
//!
//! Each child module owns one wire concern: fixed-width atoms, object
//! containers, the ordered `ExObjList` snapshot, strings, or common record
//! framing.  The public parse/serialize facade remains on the model types.

mod atoms;
mod collection;
mod containers;
mod strings;
mod wire;

pub(crate) use wire::corrupted;

#[cfg(test)]
pub(crate) use strings::encode_ole_string;
#[cfg(test)]
pub(crate) use wire::{record_bytes, record_bytes_raw};
