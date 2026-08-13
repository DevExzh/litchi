//! RTF parser owner.
//!
//! Token-to-document assembly lives in [`codec`], while focused parser
//! regression coverage lives in [`tests`].

mod codec;

#[cfg(test)]
mod tests;

pub(crate) use codec::{
    MAX_GROUP_NESTING_DEPTH, MAX_OBJECT_DATA_BYTES, MAX_OBJECTS, MAX_PICTURE_DATA_BYTES, Parser,
};
