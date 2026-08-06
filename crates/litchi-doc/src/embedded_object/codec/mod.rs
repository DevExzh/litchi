//! Wire layers for the DOC embedded-object owner.
//!
//! Parsing, encoding, low-level stream primitives, and structural validation
//! are kept separate from the snapshot transaction and OLE storage boundary.

mod encode;
mod parse;
mod primitives;
mod validation;

pub(in crate::embedded_object) use encode::{
    append_table_block, object_preview_sprms, object_separator_sprms, serialize_clx,
    serialize_fields,
};
pub(in crate::embedded_object) use parse::{managed_objects, parse_bte, parse_clx, parse_fields};
pub(in crate::embedded_object) use primitives::{
    CLX, FIB_CCP_TEXT, FIB_FC_LCB, MAX_FIELDS, MAX_PICF, MAX_PIECES, OBJ_INFO_STREAM,
    ODTPERSIST1_MUST_BE_ZERO, ODTPERSIST1_RESERVED, ODTPERSIST2_MUST_BE_ZERO, ODTPERSIST2_RESERVED,
    PLCFBTE_CHPX, PLCFFLD_MOM, SPRM_C_F_OBJ, SPRM_C_F_OLE2, SPRM_C_F_SPEC, SPRM_C_PIC_LOCATION,
    align2, align4, align512, array_at, bit, corrupted, fib_pair, put_fib_pair, put_u32, slice,
    u16_at, u32_at, word,
};
pub(in crate::embedded_object) use validation::{validate_existing_fields, validate_options};

pub(in crate::embedded_object) use super::storage::discover_targets;
